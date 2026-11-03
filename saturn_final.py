import requests
import pandas as pd
from datetime import datetime, timedelta
import os
import sys
import re
import time
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------- API клиент ----------
class SaturnClient:
    def __init__(self, base_url="https://api.fgis-saturn.ru/probeInnerArm/innerArm/seapi/"):
        self.base_url = base_url
        self.session = requests.Session()
        self.session.headers.update({"Content-Type": "application/json"})

    def _post(self, body, token, max_retries=5):
        headers = {"Authorization": f"Bearer {token}"}
        for attempt in range(max_retries):
            try:
                resp = self.session.post(self.base_url, json=body, headers=headers, timeout=60)
            except requests.exceptions.RequestException as e:
                print(f"  Ошибка сети: {e}")
                return None, 0
            if resp.status_code == 401:
                return None, 401
            if resp.status_code == 429:
                retry_after = resp.headers.get("Retry-After")
                try:
                    wait = int(retry_after) if retry_after else 15
                except ValueError:
                    wait = 15
                print(f"  ⏳ Превышен лимит запросов (429). Пауза {wait} сек. (попытка {attempt + 1}/{max_retries})...")
                time.sleep(wait)
                continue
            try:
                resp.raise_for_status()
            except requests.exceptions.HTTPError:
                return None, resp.status_code
            data = resp.json()
            if data.get("resCode") != 200:
                return None, data.get("resCode")
            return data.get("resData"), 200
        raise Exception(f"Превышен лимит запросов к API (429). Все {max_retries} попыток не удались.")

    def refresh_token(self, refresh_token):
        url = "https://fgis-saturn.ru/backend/role/newaccesstoken"
        resp = requests.get(url, headers={"refresh-token": refresh_token}, timeout=30)
        if resp.status_code == 200:
            data = resp.json()
            return data.get("access_token"), data.get("refresh_token")
        return None, None

    def get_all_invoices(self, token):
        """Загружает все накладные без фильтра по получателю"""
        all_cards = []
        pos = 0
        while True:
            body = {
                "com": "execOperation",
                "otype": "Invoice",
                "op": "static/getList()",
                "opargs": {
                    "pos": pos,
                    "size": 100,
                    "getFullCards": 1,
                    "filters": []
                }
            }
            data, status = self._post(body, token)
            if status != 200 or not data:
                break
            cards = data.get("objList", {}).get("_OBJ_ARRAY", [])
            if not cards:
                break
            all_cards.extend(cards)
            print(f"  Загружено {len(cards)} накладных (всего {len(all_cards)})")
            if len(cards) < 100:
                break
            pos += 100
            time.sleep(0.3)
        return all_cards

    def get_warehouse_by_id(self, token, warehouse_id):
        body = {
            "com": "execOperation",
            "otype": "Warehouse",
            "oid": warehouse_id,
            "op": "any/getFullCard()"
        }
        data, status = self._post(body, token)
        if status == 200 and data:
            cards = data.get("objList", {}).get("_OBJ_ARRAY", [])
            if cards:
                return cards[0].get("name", "")
        return None


# ---------- Форматирование дат ----------
def parse_date(value):
    if not value:
        return None
    try:
        if isinstance(value, (int, float)) or (isinstance(value, str) and value.isdigit()):
            return datetime.fromtimestamp(int(value))
        s = str(value).replace('Z', '+00:00')
        return pd.to_datetime(s).to_pydatetime().replace(tzinfo=None)
    except:
        return None


# ---------- Статусы ----------
STATUS_MAP = {
    "onWay": "В пути",
    "draft": "Черновик",
    "delivered": "Доставлено",
    "actual": "Актуально",
    "onDestroy": "Уничтожение",
    "destroyed": "Уничтожено",
    "onReturning": "Возврат",
    "returned": "Возвращено",
    "lostOnWay": "Потеряно в пути",
    "onPartialReturning": "Частичный возврат",
    "partialReturned": "Частично возвращено",
    "onRetailSale": "Розничная продажа",
    "onIndustrSale": "Промышленная продажа",
}


# ---------- Чтение организаций ----------
def read_organizations(excel_path):
    wb = load_workbook(excel_path, keep_vba=True)
    ws = wb['Организации']
    orgs = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[1] if len(row) > 1 else None        # B
        access = row[2] if len(row) > 2 else None       # C
        refresh = row[3] if len(row) > 3 else None      # D
        inn = str(row[4]).strip() if len(row) > 4 and row[4] else None  # E - ИНН
        if name and access:
            orgs.append({
                'name': name,
                'access_token': access,
                'refresh_token': refresh,
                'inn': inn
            })
    wb.close()
    return orgs


# ---------- Чтение настроек ----------
def read_settings(excel_path):
    wb = load_workbook(excel_path, keep_vba=True)
    ws = wb['Настройки']
    date_from = ws['B2'].value
    date_to = ws['C2'].value
    wb.close()
    return date_from, date_to


# ---------- Основная логика ----------
def main(excel_path):
    t_start = time.time()
    print(f"Чтение из {excel_path}")

    orgs = read_organizations(excel_path)
    print(f"Организаций: {len(orgs)}")

    date_from_setting, date_to_setting = read_settings(excel_path)
    print(f"Диапазон дат из настроек: {date_from_setting} — {date_to_setting}")

    client = SaturnClient()
    updated_tokens = {}
    all_org_data = {}

    # Статистика для листа СВОД
    stats = {}

    for org in orgs:
        name = org['name']
        access = org['access_token']
        refresh = org['refresh_token']
        inn = org.get('inn')
        print(f"\n=== {name} ===")

        test_body = {
            "com": "execOperation",
            "otype": "Invoice",
            "op": "static/getList()",
            "opargs": {"pos": 0, "size": 1, "getFullCards": 0, "filters": []}
        }
        _, status = client._post(test_body, access)

        error_type = None  # 'token' или 'other'

        if status == 401:
            print("  Токен истёк, обновляем...")
            new_access, new_refresh = client.refresh_token(refresh)
            if new_access:
                access, refresh = new_access, new_refresh
                updated_tokens[name] = (new_access, new_refresh)
                print("  Токен обновлён успешно")
            else:
                print("  ❌ Не удалось обновить токен, пропускаем")
                stats[name] = {'total_all': 0, 'total_onway': 0, 'matched': 0, 'internal': 0,
                               'inn_match_all': 0, 'error': True, 'error_type': 'token'}
                all_org_data[name] = []
                continue
        elif status != 200:
            print(f"  ❌ Ошибка API {status}, пропускаем")
            stats[name] = {'total_all': 0, 'total_onway': 0, 'matched': 0, 'internal': 0,
                           'inn_match_all': 0, 'error': True, 'error_type': 'other'}
            all_org_data[name] = []
            continue

        print("  Загрузка всех накладных...")
        cards = client.get_all_invoices(access)
        total_all = len(cards)
        print(f"  Получено накладных: {total_all}")

        if not cards:
            stats[name] = {'total_all': 0, 'total_onway': 0, 'matched': 0, 'internal': 0,
                           'inn_match_all': 0, 'error': False, 'error_type': None}
            all_org_data[name] = []
            continue

        # Считаем статистику по ВСЕМ накладным (любой статус) — ДО отбрасывания onWay
        inn_match_all = 0
        for card in cards:
            reciever_inn = str(card.get("recieverContractorInn", "")).strip()
            if inn and reciever_inn == inn:
                inn_match_all += 1

        # Отбрасываем всё, кроме onWay
        onway_cards = [c for c in cards if c.get("lcState", "") == "onWay"]
        total_onway = len(onway_cards)

        # Собираем уникальные ID складов ТОЛЬКО по onWay, пропуская -1
        needed_warehouse_ids = set()
        for card in onway_cards:
            if not card.get("sourceWarehouseName"):
                sid = card.get("sourceWarehouseId")
                if sid and sid != -1:
                    needed_warehouse_ids.add(sid)
            if not card.get("destinationWarehouseName"):
                did = card.get("destinationWarehauseId")
                if did and did != -1:
                    needed_warehouse_ids.add(did)

        warehouse_names = {}
        if needed_warehouse_ids:
            print(f"  Запрос названий для {len(needed_warehouse_ids)} складов...")
            for wh_id in needed_warehouse_ids:
                wh_name = client.get_warehouse_by_id(access, wh_id)
                if wh_name:
                    warehouse_names[wh_id] = wh_name
                else:
                    print(f"    Не найден склад с ID={wh_id}")
                time.sleep(0.2)

        # Формируем таблицу: onWay + даты + фильтр (внутренние или приходные)
        matched_count = 0
        internal_count = 0
        has_income_onway = False
        has_expense_only_onway = False

        rows = []
        for card in onway_cards:
            lc_state = card.get("lcState", "")

            # Фильтр по датам
            doc_date = parse_date(card.get("docDate", ""))
            if date_from_setting and doc_date:
                if isinstance(date_from_setting, datetime):
                    dt_from = date_from_setting
                else:
                    dt_from = pd.to_datetime(date_from_setting).to_pydatetime()
                if doc_date < dt_from:
                    continue
            if date_to_setting and doc_date:
                if isinstance(date_to_setting, datetime):
                    dt_to = date_to_setting + timedelta(days=1)
                else:
                    dt_to = pd.to_datetime(date_to_setting).to_pydatetime() + timedelta(days=1)
                if doc_date >= dt_to:
                    continue

            do_not_change = card.get("doNotChangeProprietor", False)
            reciever_inn = str(card.get("recieverContractorInn", "")).strip()

            # Проверяем, есть ли вообще приходные и расходные
            if inn and reciever_inn == inn:
                has_income_onway = True
            elif inn and reciever_inn != inn and reciever_inn:
                has_expense_only_onway = True

            # Фильтр для выгрузки: внутреннее перемещение ИЛИ приходная (ИНН совпадает)
            if do_not_change:
                internal_count += 1
            elif inn and reciever_inn == inn:
                matched_count += 1
            else:
                continue

            source_wh_name = card.get("sourceWarehouseName", "")
            source_id = card.get("sourceWarehouseId")
            if not source_wh_name and source_id:
                source_wh_name = warehouse_names.get(source_id, "")
            source_missing = not source_wh_name and source_id

            dest_wh_name = card.get("destinationWarehouseName", "")
            dest_id = card.get("destinationWarehauseId")
            if not dest_wh_name and dest_id:
                dest_wh_name = warehouse_names.get(dest_id, "")
            dest_missing = not dest_wh_name and dest_id

            date_recieve = parse_date(card.get("dateRecieve", ""))

            row = {
                "Код": card.get("id", ""),
                "Статус": STATUS_MAP.get(lc_state, lc_state),
                "Дата регистрации": doc_date,
                "№ накладной": card.get("docNum", ""),
                "Наименование": card.get("name", ""),
                "Склад-источник": source_wh_name,
                "Склад-приемник": dest_wh_name,
                "Дата получения": date_recieve,
                "Номер ТС": card.get("vehicleNumber", ""),
                "Отправитель": card.get("ownerContractorName", ""),
                "ИНН получателя": reciever_inn,
                "Действия": "",
                "_source_missing": source_missing,
                "_dest_missing": dest_missing,
            }
            rows.append(row)

        all_org_data[name] = rows

        # Определяем статус
        if total_onway == 0:
            status_text = "нет накладных в пути"
        elif has_income_onway:
            status_text = "есть приходные в пути"
        elif has_expense_only_onway:
            status_text = "есть только расходные в пути"
        else:
            status_text = "нет накладных в пути"

        stats[name] = {
            'total_all': total_all,
            'total_onway': total_onway,
            'matched': matched_count,
            'internal': internal_count,
            'inn_match_all': inn_match_all,
            'error': False,
            'error_type': None,
            'status_text': status_text
        }
        print(f"  Всего: {total_all}, в пути: {total_onway}, приходных: {matched_count}, внутр.: {internal_count}, ИНН совпало всего: {inn_match_all}")

    # Сохраняем результат
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    out_path = os.path.join(os.path.dirname(excel_path), f'Saturn_Report_{timestamp}.xlsx')
    print(f"\nСохранение отчёта в {out_path}")

    # Цвета
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    green_font = Font(color="006400", bold=True)
    red_font = Font(color="FF0000", bold=True)
    wrap_alignment = Alignment(wrap_text=True, vertical="top")
    light_green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    bright_red_fill = PatternFill(start_color="FF4444", end_color="FF4444", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    column_widths = {
        "Код": 12,
        "Статус": 14,
        "Дата регистрации": 22.4,
        "№ накладной": 28,
        "Наименование": 28,
        "Склад-источник": 35,
        "Склад-приемник": 35,
        "Дата получения": 22.4,
        "Номер ТС": 19.6,
        "Отправитель": 28,
        "ИНН получателя": 18,
        "Действия": 14,
    }

    columns_order = ["Код", "Статус", "Дата регистрации", "№ накладной", "Наименование",
                     "Склад-источник", "Склад-приемник", "Дата получения", "Номер ТС",
                     "Отправитель", "ИНН получателя", "Действия"]

    has_any_data = any(len(rows) > 0 for rows in all_org_data.values())

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        # ====== Лист СВОД ======
        svod_rows = []
        for org in orgs:
            name = org['name']
            st = stats.get(name, {
                'total_all': 0, 'total_onway': 0, 'matched': 0, 'internal': 0,
                'inn_match_all': 0, 'error': True, 'error_type': 'token', 'status_text': 'не получилось выгрузить накладные из-за проблем с ключами'
            })

            total_all = st['total_all']
            if st['error']:
                if st.get('error_type') == 'token':
                    status_text = "не получилось выгрузить накладные из-за проблем с ключами"
                else:
                    status_text = "не получилось выгрузить накладные из-за других проблем"
                pct = None
            else:
                status_text = st.get('status_text', 'нет накладных в пути')
                if total_all > 0:
                    pct = round(st['inn_match_all'] / total_all * 100, 1)
                else:
                    pct = 0

            svod_rows.append({
                "Организация": name,
                "Всего накладных": total_all,
                "Из них в пути": st['total_onway'],
                "Приходных": st['matched'],
                "Внутр.перемещений": st['internal'],
                "Итого выгружено": st['matched'] + st['internal'],
                "% совпадений ИНН от всех накладных": pct if pct is not None else "ОШИБКА",
                "Статус": status_text
            })

        df_svod = pd.DataFrame(svod_rows)
        df_svod.to_excel(writer, sheet_name='СВОД', index=False, startrow=0)
        ws_svod = writer.sheets['СВОД']

        for col_idx in range(1, len(df_svod.columns) + 1):
            cell = ws_svod.cell(row=1, column=col_idx)
            cell.font = Font(color="006400", bold=True)
            cell.alignment = wrap_alignment
            cell.border = thin_border
            ws_svod.column_dimensions[get_column_letter(col_idx)].width = 24

        for row_idx in range(2, len(df_svod) + 2):
            for col_idx in range(1, len(df_svod.columns) + 1):
                cell = ws_svod.cell(row=row_idx, column=col_idx)
                cell.alignment = wrap_alignment
                cell.border = thin_border

            org_name_cell = ws_svod.cell(row=row_idx, column=1).value
            st = stats.get(org_name_cell, {})

            if st.get('error'):
                for col_idx in range(1, len(df_svod.columns) + 1):
                    ws_svod.cell(row=row_idx, column=col_idx).fill = bright_red_fill
            elif st.get('matched', 0) + st.get('internal', 0) > 0:
                for col_idx in range(1, len(df_svod.columns) + 1):
                    ws_svod.cell(row=row_idx, column=col_idx).fill = light_green_fill

        print("  Лист СВОД создан")

        # ====== Листы организаций ======
        if not has_any_data:
            pd.DataFrame({"": ["Нет данных ни по одной организации"]}).to_excel(
                writer, sheet_name='Результат', index=False, header=False)
            ws = writer.sheets['Результат']
            ws.cell(row=1, column=1).font = red_font
            ws.column_dimensions['A'].width = 50
        else:
            for org, rows in all_org_data.items():
                sheet_name = re.sub(r'[\\/*?\[\]]', '', str(org))[:31]

                if not rows:
                    df_empty = pd.DataFrame({"": ["Нет приходных накладных в пути за этот период"]})
                    df_empty.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=0)
                    ws = writer.sheets[sheet_name]
                    ws.cell(row=1, column=1).font = red_font
                    ws.column_dimensions['A'].width = 40
                    print(f"  {org}: нет данных")
                    continue

                clean_rows = []
                for r in rows:
                    clean_row = {k: v for k, v in r.items() if not k.startswith("_")}
                    clean_rows.append(clean_row)

                df = pd.DataFrame(clean_rows)
                df = df[columns_order]
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

                ws = writer.sheets[sheet_name]

                for col_idx in range(1, len(columns_order) + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.font = green_font
                    cell.alignment = wrap_alignment
                    cell.border = thin_border

                for col_idx, col_name in enumerate(columns_order, 1):
                    col_letter = get_column_letter(col_idx)
                    ws.column_dimensions[col_letter].width = column_widths.get(col_name, 14)

                source_col = columns_order.index("Склад-источник") + 1
                dest_col = columns_order.index("Склад-приемник") + 1
                date_reg_col = columns_order.index("Дата регистрации") + 1
                date_rec_col = columns_order.index("Дата получения") + 1
                code_col = columns_order.index("Код") + 1

                for row_idx, row_data in enumerate(rows, 2):
                    for col_idx in range(1, len(columns_order) + 1):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.alignment = wrap_alignment
                        cell.border = thin_border

                    if row_data.get("_source_missing"):
                        ws.cell(row=row_idx, column=source_col).fill = yellow_fill
                    if row_data.get("_dest_missing"):
                        ws.cell(row=row_idx, column=dest_col).fill = yellow_fill

                    code_cell = ws.cell(row=row_idx, column=code_col)
                    if isinstance(code_cell.value, (int, float)):
                        code_cell.number_format = '0'

                    reg_cell = ws.cell(row=row_idx, column=date_reg_col)
                    if reg_cell.value and isinstance(reg_cell.value, datetime):
                        reg_cell.number_format = 'DD.MM.YYYY'

                    rec_cell = ws.cell(row=row_idx, column=date_rec_col)
                    if rec_cell.value and isinstance(rec_cell.value, datetime):
                        rec_cell.number_format = 'DD.MM.YYYY'

                print(f"  {org}: {len(clean_rows)} строк")

    # Обновляем токены в столбцах C и D
    if updated_tokens:
        print("\nОбновление токенов в исходном файле...")
        wb = load_workbook(excel_path, keep_vba=True)
        ws = wb['Организации']

        for row_idx in range(2, ws.max_row + 1):
            org_name_cell = ws.cell(row=row_idx, column=2)
            if org_name_cell.value in updated_tokens:
                new_access, new_refresh = updated_tokens[org_name_cell.value]
                ws.cell(row=row_idx, column=3).value = new_access
                ws.cell(row=row_idx, column=4).value = new_refresh
                print(f"  Обновлены токены для: {org_name_cell.value}")

        wb.save(excel_path)
        wb.close()
        print("  Токены сохранены в исходный файл")

    # Финальное сообщение с временем выполнения
    elapsed = time.time() - t_start
    minutes = int(elapsed // 60)
    seconds = int(elapsed % 60)
    elapsed_str = f"{minutes} мин {seconds} сек" if minutes else f"{seconds} сек"

    print(f"\nГотово за {elapsed_str}!")
    try:
        import tkinter.messagebox as mb
        mb.showinfo("Готово", f"Отчёт сохранён:\n{out_path}\n\nВремя выполнения: {elapsed_str}")
    except:
        print(f"\nОтчёт сохранён в {out_path}")
        input("Нажмите Enter для выхода...")

    return out_path


if __name__ == '__main__':
    import warnings
    warnings.filterwarnings("ignore")

    def get_program_dir():
        """Возвращает папку, где лежит программа (.exe / бинарник / .app / .py)."""
        # Запущено как собранный файл
        if getattr(sys, 'frozen', False):
            exe_path = os.path.abspath(sys.executable)
            # macOS .app bundle: .../SaturnReport.app/Contents/MacOS/SaturnReport
            # Возвращаем папку, где лежит сам .app
            if ".app/Contents/MacOS" in exe_path:
                app_parent = exe_path.split(".app/Contents/MacOS")[0]
                return os.path.dirname(app_parent) or os.sep
            # Windows .exe / Linux бинарник — папка бинарника
            return os.path.dirname(exe_path)

        # Запущено как .py скрипт
        return os.path.dirname(os.path.abspath(__file__))

    if len(sys.argv) < 2:
        program_dir = get_program_dir()
        excel_path = os.path.join(program_dir, "Saturn_Runner.xlsm")
        print(f"Папка программы: {program_dir}")
        print(f"Ищем файл настроек: {excel_path}")

        if not os.path.exists(excel_path):
            err_text = (
                f"Файл Saturn_Runner.xlsm не найден.\n\n"
                f"Ожидаемый путь:\n{excel_path}\n\n"
                f"Положите Saturn_Runner.xlsm в ту же папку, где находится программа."
            )
            try:
                import tkinter.messagebox as mb
                mb.showerror("Файл не найден", err_text)
            except:
                print(err_text)
                input("Нажмите Enter для выхода...")
            sys.exit(1)
    else:
        excel_path = sys.argv[1]

    print(f"Используем файл настроек: {excel_path}")

    try:
        main(excel_path)
    except Exception as e:
        import traceback
        err_text = f"{type(e).__name__}: {e}\n\n{traceback.format_exc()}"
        try:
            import tkinter.messagebox as mb
            mb.showerror("Ошибка выполнения", err_text)
        except:
            print(err_text)
            input("Нажмите Enter для выхода...")
