# -*- coding: utf-8 -*-
"""B0: probe ручных Документ.СписаниеТоваров за январь–май 2026.

v4: использует ПРЕДСТАВЛЕНИЕ() в запросе для всех ссылок 1С.
Группировка по комментариям. Подробная статистика товаров.
"""
import io
import json
import os
import sys
from collections import Counter, defaultdict

import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _tl_config import BASE_PATH, USER, PWD  # noqa: E402

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

OUT_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "probe_writeoff_irina_result.json")


_repr_cache = {}
_repr_conn = None


def repr_ref(val):
    """Получить ПРЕДСТАВЛЕНИЕ ссылки 1С через одиночный запрос."""
    if val is None:
        return ""
    s = str(val)
    if not s.startswith("<COMObject"):
        return s
    key = id(val)
    if key in _repr_cache:
        return _repr_cache[key]
    try:
        q = _repr_conn.NewObject("Запрос")
        q.УстановитьПараметр("Объект", val)
        q.Текст = "ВЫБРАТЬ ПРЕДСТАВЛЕНИЕ(&Объект) КАК Имя"
        sel = q.Выполнить().Выбрать()
        if sel.Следующий():
            _repr_cache[key] = str(sel.Имя)
            return _repr_cache[key]
    except Exception:
        pass
    return s


def main():
    global _repr_conn
    com = win32com.client.Dispatch("V83.COMConnector")
    conn = com.Connect(f'File="{BASE_PATH}";Usr="{USER}";Pwd="{PWD}";')
    _repr_conn = conn

    print("=" * 95)
    print("B0: Документ.СписаниеТоваров за январь–май 2026 (стенд GIG Base2)")
    print("=" * 95)

    # ---- 1. Шапки документов с ПРЕДСТАВЛЕНИЯМИ всех ссылок ----
    q = conn.NewObject("Запрос")
    q.Текст = """
    ВЫБРАТЬ
        Док.Ссылка                              КАК Ссылка,
        Док.Номер                               КАК Номер,
        Док.Дата                                КАК Дата,
        ПРЕДСТАВЛЕНИЕ(Док.Организация)          КАК Организация,
        ПРЕДСТАВЛЕНИЕ(Док.Склад)                КАК Склад,
        ПРЕДСТАВЛЕНИЕ(Док.ВидОперации)          КАК ВидОперации,
        ПРЕДСТАВЛЕНИЕ(Док.Ответственный)        КАК Ответственный,
        Док.Комментарий                         КАК Комментарий,
        Док.Проведен                            КАК Проведен
    ИЗ
        Документ.СписаниеТоваров КАК Док
    ГДЕ
        Док.Дата >= ДАТАВРЕМЯ(2026, 1, 1)
        И Док.Дата <  ДАТАВРЕМЯ(2026, 6, 1)
    УПОРЯДОЧИТЬ ПО Док.Дата
    """
    sel = q.Выполнить().Выбрать()

    headers = []
    while sel.Следующий():
        headers.append({
            "_ref":         sel.Ссылка,
            "Номер":        str(sel.Номер),
            "Дата":         str(sel.Дата),
            "Организация":  str(sel.Организация),
            "Склад":        str(sel.Склад),
            "ВидОперации":  str(sel.ВидОперации),
            "Ответственный": str(sel.Ответственный),
            "Комментарий":  str(sel.Комментарий),
            "Проведен":     bool(sel.Проведен),
        })

    print(f"Найдено документов: {len(headers)}")
    if not headers:
        print()
        print("Документов нет за период.")
        return

    # ---- 2. Сводка ----
    by_op  = Counter(h["ВидОперации"] for h in headers)
    by_stock = Counter(h["Склад"] for h in headers)
    by_user  = Counter(h["Ответственный"] for h in headers)
    by_com   = Counter(h["Комментарий"].strip().lower() or "(пусто)" for h in headers)

    print()
    print("По ВидОперации:")
    for k, v in by_op.most_common():
        print(f"  {k:60} {v:>4}")
    print()
    print("По Складу:")
    for k, v in by_stock.most_common():
        print(f"  {k:60} {v:>4}")
    print()
    print("По Ответственному:")
    for k, v in by_user.most_common():
        print(f"  {k:60} {v:>4}")
    print()
    print("По Комментарию (нормализованному):")
    for k, v in by_com.most_common():
        print(f"  {k:60} {v:>4}")

    # ---- 3. Детали по каждому документу ----
    details = []
    for i, h in enumerate(headers, 1):
        ref = h["_ref"]

        # 3.1 ТЧ Товары — через ПолучитьОбъект (надёжнее ВТ-запроса)
        tovary = []
        try:
            obj = ref.ПолучитьОбъект()
            for row in obj.Товары:
                item = {}
                # Перебираем стандартные имена реквизитов ТЧ Товары
                for fld in ("Номенклатура", "Количество", "СчетУчета",
                            "СтранаПроисхождения", "НомерГТД",
                            "Цена", "Сумма", "СтавкаНДС", "СуммаНДС",
                            "ХарактеристикаНоменклатуры", "Партия",
                            "СчетЗатрат", "СтатьяЗатрат", "ПодразделениеЗатрат",
                            "НоменклатурнаяГруппа", "СтатьиЗатрат"):
                    try:
                        if hasattr(row, fld):
                            val = getattr(row, fld)
                            if val is not None:
                                if isinstance(val, (int, float)):
                                    item[fld] = float(val)
                                else:
                                    item[fld] = repr_ref(val)
                    except Exception:
                        pass
                tovary.append(item)
        except Exception as e:
            tovary = [{"_error": str(e)}]

        # 3.2 Проводки регистра бухгалтерии
        q_acc = conn.NewObject("Запрос")
        q_acc.УстановитьПараметр("Док", ref)
        q_acc.Текст = """
        ВЫБРАТЬ
            ПРЕДСТАВЛЕНИЕ(Дв.СчетДт)        КАК СчетДт,
            ПРЕДСТАВЛЕНИЕ(Дв.СчетКт)        КАК СчетКт,
            Дв.Сумма                         КАК Сумма,
            Дв.КоличествоДт                  КАК КолДт,
            Дв.КоличествоКт                  КАК КолКт,
            ПРЕДСТАВЛЕНИЕ(Дв.СубконтоДт1)   КАК СубДт1,
            ПРЕДСТАВЛЕНИЕ(Дв.СубконтоДт2)   КАК СубДт2,
            ПРЕДСТАВЛЕНИЕ(Дв.СубконтоДт3)   КАК СубДт3,
            ПРЕДСТАВЛЕНИЕ(Дв.СубконтоКт1)   КАК СубКт1,
            ПРЕДСТАВЛЕНИЕ(Дв.СубконтоКт2)   КАК СубКт2,
            ПРЕДСТАВЛЕНИЕ(Дв.СубконтоКт3)   КАК СубКт3,
            Дв.Содержание                    КАК Содержание
        ИЗ РегистрБухгалтерии.Хозрасчетный КАК Дв
        ГДЕ Дв.Регистратор = &Док
        УПОРЯДОЧИТЬ ПО Дв.НомерСтроки
        """
        provodki = []
        try:
            s = q_acc.Выполнить().Выбрать()
            while s.Следующий():
                provodki.append({
                    "СчетДт":     str(s.СчетДт),
                    "СчетКт":     str(s.СчетКт),
                    "Сумма":      float(s.Сумма) if s.Сумма else 0,
                    "КолДт":      float(s.КолДт) if s.КолДт else 0,
                    "КолКт":      float(s.КолКт) if s.КолКт else 0,
                    "СубДт1":     str(s.СубДт1),
                    "СубДт2":     str(s.СубДт2),
                    "СубДт3":     str(s.СубДт3),
                    "СубКт1":     str(s.СубКт1),
                    "СубКт2":     str(s.СубКт2),
                    "СубКт3":     str(s.СубКт3),
                    "Содержание": str(s.Содержание),
                })
        except Exception as e:
            provodki = [{"_error": str(e)}]

        # 3.3 ТоварыНаАЗК (наш ключевой регистр для топлива)
        tovary_na_azk = []
        try:
            q_reg = conn.NewObject("Запрос")
            q_reg.УстановитьПараметр("Док", ref)
            q_reg.Текст = """
            ВЫБРАТЬ
                ПРЕДСТАВЛЕНИЕ(Дв.Номенклатура)  КАК Номенклатура,
                ПРЕДСТАВЛЕНИЕ(Дв.Склад)          КАК Склад,
                Дв.Количество                    КАК Количество,
                ПРЕДСТАВЛЕНИЕ(Дв.ВидДвижения)   КАК ВидДвижения
            ИЗ РегистрНакопления.ТоварыНаАЗК КАК Дв
            ГДЕ Дв.Регистратор = &Док
            """
            s = q_reg.Выполнить().Выбрать()
            while s.Следующий():
                tovary_na_azk.append({
                    "Номенклатура": str(s.Номенклатура),
                    "Склад":        str(s.Склад),
                    "Количество":   float(s.Количество) if s.Количество else 0,
                    "ВидДвижения":  str(s.ВидДвижения),
                })
        except Exception:
            pass  # регистра может не быть

        # Аналогично ТоварыОрганизаций (если БП ведёт)
        tovary_org = []
        try:
            q_reg = conn.NewObject("Запрос")
            q_reg.УстановитьПараметр("Док", ref)
            q_reg.Текст = """
            ВЫБРАТЬ
                ПРЕДСТАВЛЕНИЕ(Дв.Номенклатура)   КАК Номенклатура,
                ПРЕДСТАВЛЕНИЕ(Дв.Склад)           КАК Склад,
                ПРЕДСТАВЛЕНИЕ(Дв.Организация)    КАК Организация,
                Дв.Количество                     КАК Количество,
                ПРЕДСТАВЛЕНИЕ(Дв.ВидДвижения)    КАК ВидДвижения
            ИЗ РегистрНакопления.ТоварыОрганизаций КАК Дв
            ГДЕ Дв.Регистратор = &Док
            """
            s = q_reg.Выполнить().Выбрать()
            while s.Следующий():
                tovary_org.append({
                    "Номенклатура": str(s.Номенклатура),
                    "Склад":        str(s.Склад),
                    "Организация":  str(s.Организация),
                    "Количество":   float(s.Количество) if s.Количество else 0,
                    "ВидДвижения":  str(s.ВидДвижения),
                })
        except Exception:
            pass

        h_clean = {k: v for k, v in h.items() if k != "_ref"}
        details.append({
            "header":           h_clean,
            "tovary":           tovary,
            "provodki":         provodki,
            "tovary_na_azk":    tovary_na_azk,
            "tovary_org":       tovary_org,
        })

    # ---- 4. Группировка по типу комментария — что НОВОЕ интересно ----
    print()
    print("=" * 95)
    print("ОБРАЗЦЫ документов по категории комментария:")
    print("=" * 95)

    by_category = defaultdict(list)
    for d in details:
        cat = d["header"]["Комментарий"].strip().lower() or "(пусто)"
        by_category[cat].append(d)

    for cat, docs in sorted(by_category.items(), key=lambda x: -len(x[1])):
        print()
        print(f"--- Категория «{cat}» — {len(docs)} документов")
        sample = docs[0]
        h = sample["header"]
        print(f"    Пример: {h['Номер']} от {h['Дата']}")
        print(f"    Склад:          {h['Склад']}")
        print(f"    ВидОперации:    {h['ВидОперации']}")
        print(f"    Ответственный:  {h['Ответственный']}")
        print(f"    Проведен:       {h['Проведен']}")
        print(f"    ТЧ Товары ({len(sample['tovary'])}):")
        for t in sample["tovary"][:8]:
            nom = str(t.get('Номенклатура',''))[:50]
            kol = t.get('Количество','')
            sch = str(t.get('СчетУчета',''))
            sz  = str(t.get('СчетЗатрат',''))
            sa  = str(t.get('СтатьяЗатрат',''))
            ng  = str(t.get('НоменклатурнаяГруппа',''))
            print(f"      {nom:50} кол={kol:>10}  счУ={sch:8}  счЗ={sz:6}  стЗ={sa:25}  НГ={ng}")
        if len(sample["tovary"]) > 8:
            print(f"      ... ещё {len(sample['tovary']) - 8}")
        # Проводки убраны — на стенде пустые. Боевая БП покажет
        print(f"    ТоварыНаАЗК ({len(sample['tovary_na_azk'])}):")
        for r in sample["tovary_na_azk"][:3]:
            print(f"      {r.get('Номенклатура','')[:40]:40} "
                  f"кол={r.get('Количество',''):>10}  "
                  f"скл={r.get('Склад','')}")
        print(f"    ТоварыОрганизаций ({len(sample['tovary_org'])}):")
        for r in sample["tovary_org"][:3]:
            print(f"      {r.get('Номенклатура','')[:40]:40} "
                  f"кол={r.get('Количество',''):>10}  "
                  f"скл={r.get('Склад','')}")

    # ---- 5. Сохранение JSON ----
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump({
            "base":         BASE_PATH,
            "period":       "2026-01 … 2026-05",
            "total":        len(headers),
            "by_operation": dict(by_op),
            "by_stock":     dict(by_stock),
            "by_user":      dict(by_user),
            "by_comment":   dict(by_com),
            "documents":    details,
        }, f, ensure_ascii=False, indent=2, default=str)
    print()
    print("=" * 95)
    print(f"Полные данные → {OUT_JSON}")


if __name__ == "__main__":
    main()
