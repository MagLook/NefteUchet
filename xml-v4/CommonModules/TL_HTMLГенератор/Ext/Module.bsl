////////////////////////////////////////////////////////////////////////////////
// Общий модуль TL_HTMLГенератор
// Расширение ElsyPlus Ledger для 1С:БП 3.0
// Контекст: Сервер, ВызовСервера
// Формирование HTML для дашборда и панели деталей.
//
// Цветовая схема (единая с TradeFrame):
//   Primary #3B82F6 — акцент, заголовки (hsl 217 91% 60%)
//   Success #16A34A — проведён, OK (hsl 142 76% 36%)
//   Warning #EAB308 — загружен, ожидание (hsl 45 93% 47%)
//   Error   #EF4444 — ошибка (hsl 0 84% 60%)
//   Purple  #A855F7 — ТТН (hsl 280 80% 60%)
//   Muted   #6B7280 — второстепенный текст
//   BG      #F0F5FA — фон (hsl 220 14% 94%)
//   Card    #FFFFFF — карточки
//   Border  #D9DFE7 — границы (hsl 220 13% 87%)
//   Shadow  0 1px 3px rgba(0,0,0,0.08) — мягкие тени
////////////////////////////////////////////////////////////////////////////////

#Область ПрограммныйИнтерфейс

// Сформировать HTML дашборда (верхняя полоса с 5 карточками).
//
// Параметры:
//   ВсегоСмен     - Число
//   Загружено     - Число - документов в 1С
//   Новых         - Число - ещё не загружено
//   Ошибок        - Число
//   СуммаЗаПериод - Число
//
// Возвращаемое значение:
//   Строка - полный HTML документ
//
Функция СформироватьДашборд(ВсегоСмен, Загружено, Новых, Ошибок, СуммаЗаПериод,
	СтатусSTS = "", ТекстЛога = "", КоличествоТТН = 0, ИнфоСтанция = "", ИнфоОстатки = "") Экспорт

	СуммаТекст = ФорматСуммы(СуммаЗаПериод);

	// === Контейнер: 4 колонки в одну строку ===
	СтильПанели = "background:#fff; border:1px solid #E2E8F0; border-radius:6px; padding:6px 10px; overflow:hidden;";

	HTML = "<!DOCTYPE html><html><head><meta charset=""utf-8""></head>"
		+ "<body style=""margin:0; padding:6px 8px; font-family:-apple-system,'Segoe UI',sans-serif;"
		+ " background:#F0F4F8; font-size:12px;"">"
		+ "<div style=""display:flex; gap:6px; height:100%;"">";

	// === КОЛОНКА 1: KPI (25%) — сетка 3x2, растянуть на всю область ===
	HTML = HTML + "<div style=""width:25%; " + СтильПанели + " display:flex; flex-direction:column;"">"
		+ "<div style=""font-size:10px; color:#94A3B8; font-weight:600; margin-bottom:5px; text-transform:uppercase; letter-spacing:0.5px;"">KPI</div>"
		+ "<div style=""display:grid; grid-template-columns:1fr 1fr 1fr; grid-template-rows:1fr 1fr; gap:4px; flex:1;"">"
		+ МиниКарточкаКомпакт(XMLСтрока(ВсегоСмен), "Смен", "#3B82F6")
		+ МиниКарточкаКомпакт(XMLСтрока(Загружено), "В 1С", "#16A34A")
		+ МиниКарточкаКомпакт(XMLСтрока(Новых), "Новых", "#EAB308")
		+ МиниКарточкаКомпакт(XMLСтрока(КоличествоТТН), "ТТН", "#8B5CF6")
		+ МиниКарточкаКомпакт(XMLСтрока(Ошибок), "Ошибок", "#EF4444")
		+ МиниКарточкаКомпакт(СуммаТекст, "Выруч.", "#6B7280")
		+ "</div></div>";

	// === КОЛОНКА 2: STS (20%) ===
	ЦветSTS = ?(СтрНайти(НРег(СтатусSTS), "подключен") > 0, "#16A34A", "#EF4444");
	HTML = HTML + "<div style=""width:20%; " + СтильПанели + """>"
		+ "<div style=""font-size:10px; color:#94A3B8; font-weight:600; margin-bottom:4px; text-transform:uppercase; letter-spacing:0.5px;"">STS API</div>";
	Если ЗначениеЗаполнено(СтатусSTS) Тогда
		HTML = HTML + "<div style=""margin-bottom:3px;"">"
			+ "<span style=""color:" + ЦветSTS + "; font-size:14px;"">&#9679;</span> "
			+ "<span style=""color:#475569; font-size:11px;"">" + Экр(СтатусSTS) + "</span></div>";
	КонецЕсли;
	Если ЗначениеЗаполнено(ИнфоСтанция) Тогда
		HTML = HTML + "<div style=""color:#64748B; font-size:10px; line-height:1.4;"">" + Экр(ИнфоСтанция) + "</div>";
	КонецЕсли;
	HTML = HTML + "</div>";

	// === КОЛОНКА 3: 1С (25%) ===
	HTML = HTML + "<div style=""width:25%; " + СтильПанели + """>"
		+ "<div style=""font-size:10px; color:#94A3B8; font-weight:600; margin-bottom:5px; text-transform:uppercase; letter-spacing:0.5px;"">Остатки 41.02</div>";
	Если ЗначениеЗаполнено(ИнфоОстатки) Тогда
		HTML = HTML + "<div style=""font-size:11px; line-height:1.6;"">" + ИнфоОстатки + "</div>";
	Иначе
		HTML = HTML + "<div style=""color:#CBD5E1; font-size:11px;"">Нажмите Обновить</div>";
	КонецЕсли;
	HTML = HTML + "</div>";

	// === КОЛОНКА 4: Лог (30%) ===
	HTML = HTML + "<div style=""width:30%; " + СтильПанели
		+ " font-family:Consolas,'Courier New',monospace; font-size:10px;"">"
		+ "<div style=""font-size:10px; color:#94A3B8; font-weight:600; margin-bottom:4px;"
		+ " text-transform:uppercase; letter-spacing:0.5px; font-family:-apple-system,'Segoe UI',sans-serif;"">Лог</div>";
	Если ЗначениеЗаполнено(ТекстЛога) Тогда
		Строки = СтрРазделить(ТекстЛога, Символы.ПС);
		МаксСтрок = ?(Строки.Количество() > 6, 6, Строки.Количество());
		Для Инд = 0 По МаксСтрок - 1 Цикл
			Стр = Строки[Инд];
			Если Не ЗначениеЗаполнено(Стр) Тогда Продолжить; КонецЕсли;
			Если СтрНайти(НРег(Стр), "ошибка") > 0 ИЛИ СтрНайти(НРег(Стр), "error") > 0 Тогда
				ЦветЛог = "#EF4444";
			ИначеЕсли СтрНайти(НРег(Стр), "ok") > 0 ИЛИ СтрНайти(НРег(Стр), "загружен") > 0
				ИЛИ СтрНайти(НРег(Стр), "проведен") > 0 Тогда
				ЦветЛог = "#16A34A";
			Иначе
				ЦветЛог = "#94A3B8";
			КонецЕсли;
			HTML = HTML + "<div style=""color:#475569; line-height:1.4; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;"">"
				+ "<span style=""color:" + ЦветЛог + ";"">&#9679;</span> " + Экр(Стр) + "</div>";
		КонецЦикла;
	Иначе
		HTML = HTML + "<div style=""color:#CBD5E1;"">Нет записей</div>";
	КонецЕсли;
	HTML = HTML + "</div>";

	HTML = HTML + "</div></body></html>";

	Возврат HTML;

КонецФункции

// Сформировать HTML деталей для сменного отчёта.
//
// Параметры:
//   ДанныеJSON          - Строка - JSON сменного отчёта
//   Заголовок           - Строка - "Смена #143, АЗС-208, 22.03.2026"
//   СвязанныеДокументы  - Массив из Структура("ТипДокумента, Описание, Статус")
//
// Возвращаемое значение:
//   Строка - HTML
//
Функция СформироватьДеталиСмены(ДанныеJSON, Заголовок, СвязанныеДокументы = Неопределено) Экспорт

	Попытка
		Данные = TL_СозданиеДокументов.ПолучитьИзJSON(ДанныеJSON);
	Исключение
		Возврат ОбёрткаHTML("<p style=""color:#c00;"">Ошибка разбора JSON</p>");
	КонецПопытки;

	Если Данные = Неопределено Тогда
		Возврат ОбёрткаHTML("<p style=""color:#888;"">Нет данных</p>");
	КонецЕсли;

	// ===== СБОР ДАННЫХ =====

	// Коды топлива из sales (уникальные, для колонок)
	КодыТоплива = Новый Массив; // "АИ-92", "АИ-95", "ДТ"
	ИменаТоплива = Новый Соответствие;

	// Агрегат по оплатам: ИмяОплаты → Соответствие(КодТоплива → Структура(Литры, Сумма))
	АгрегатОплат = Новый Соответствие;
	ИтогиПоТопливу = Новый Соответствие; // КодТоплива → Структура(Литры, Сумма)
	ИтогоВсеЛитры = 0;
	ИтогоВсеСумма = 0;

	МассивПродаж = TL_ApiКлиент.ПолучитьЗначение(Данные, "sales", Неопределено);
	Если ТипЗнч(МассивПродаж) = Тип("Массив") Тогда
		Для Каждого Продажа Из МассивПродаж Цикл
			ТипОпл = TL_ApiКлиент.ПолучитьЗначение(Продажа, "pay_type", Неопределено);
			Если ТипОпл = Неопределено Тогда Продолжить; КонецЕсли;
			ИмяОпл = Строка(TL_ApiКлиент.ПолучитьЗначение(ТипОпл, "name", ""));

			Если АгрегатОплат.Получить(ИмяОпл) = Неопределено Тогда
				АгрегатОплат.Вставить(ИмяОпл, Новый Соответствие);
			КонецЕсли;

			МассивТопл = TL_ApiКлиент.ПолучитьЗначение(Продажа, "fuel", Неопределено);
			Если ТипЗнч(МассивТопл) <> Тип("Массив") Тогда Продолжить; КонецЕсли;

			Для Каждого ЭлТ Из МассивТопл Цикл
				Сервис = TL_ApiКлиент.ПолучитьЗначение(ЭлТ, "service", Неопределено);
				Если Сервис = Неопределено Тогда Продолжить; КонецЕсли;
				КодТ = Строка(TL_ApiКлиент.ПолучитьЗначение(Сервис, "service_name", ""));

				Если ИменаТоплива.Получить(КодТ) = Неопределено Тогда
					КодыТоплива.Добавить(КодТ);
					ИменаТоплива.Вставить(КодТ, КодТ);
					ИтогиПоТопливу.Вставить(КодТ, Новый Структура("Литры, Сумма", 0, 0));
				КонецЕсли;

				Выпуск = TL_ApiКлиент.ПолучитьЗначение(ЭлТ, "release", Неопределено);
				Если Выпуск = Неопределено Тогда Продолжить; КонецЕсли;
				Л = Число(TL_ApiКлиент.ПолучитьЗначение(Выпуск, "volume", 0));
				С = Число(TL_ApiКлиент.ПолучитьЗначение(Выпуск, "cost", 0));

				ДанныеОпл = АгрегатОплат.Получить(ИмяОпл);
				Существ = ДанныеОпл.Получить(КодТ);
				Если Существ <> Неопределено Тогда
					Существ.Литры = Существ.Литры + Л;
					Существ.Сумма = Существ.Сумма + С;
				Иначе
					ДанныеОпл.Вставить(КодТ, Новый Структура("Литры, Сумма", Л, С));
				КонецЕсли;

				Итог = ИтогиПоТопливу.Получить(КодТ);
				Итог.Литры = Итог.Литры + Л;
				Итог.Сумма = Итог.Сумма + С;
				ИтогоВсеЛитры = ИтогоВсеЛитры + Л;
				ИтогоВсеСумма = ИтогоВсеСумма + С;
			КонецЦикла;
		КонецЦикла;
	КонецЕсли;

	КолТоплива = КодыТоплива.Количество();

	// ===== ПРЕДВАРИТЕЛЬНЫЙ РАСЧЁТ: РОЗНИЦА vs БЕЗНАЛ =====
	// Розница = retail_cash + retail_card. Только у неё есть учётная сумма.
	// Все остальные каналы (VIAcard, БАЛТОП, ЯНДЕКС, Ведомости, Талоны) — только литры.
	// Расценка по ним происходит в конце месяца на виртуальных складах.

	КешЭтоРозница = Новый Соответствие; // ИмяОпл → Булево
	РозницаПоТопливу = Новый Соответствие;   // КодТ → Структура(Литры, Сумма)
	Для Каждого КодТ Из КодыТоплива Цикл
		РозницаПоТопливу.Вставить(КодТ, Новый Структура("Литры, Сумма", 0, 0));
	КонецЦикла;
	РозницаИтого = 0;

	Для Каждого КЗ Из АгрегатОплат Цикл
		ИмяОпл = КЗ.Ключ;
		ИмяНРегВр = НРег(ИмяОпл);
		ЭР = (СтрНайти(ИмяНРегВр, "наличн") > 0) ИЛИ (СтрНайти(ИмяНРегВр, "сбербанк") > 0);
		КешЭтоРозница.Вставить(ИмяОпл, ЭР);
		Если ЭР Тогда
			Для Каждого КодТ Из КодыТоплива Цикл
				Яч = КЗ.Значение.Получить(КодТ);
				Если Яч <> Неопределено И Яч.Литры > 0 Тогда
					Р = РозницаПоТопливу.Получить(КодТ);
					Р.Литры = Р.Литры + Яч.Литры;
					Р.Сумма = Р.Сумма + Яч.Сумма;
					РозницаИтого = РозницаИтого + Яч.Сумма;
				КонецЕсли;
			КонецЦикла;
		КонецЕсли;
	КонецЦикла;

	БезналЛитры = ИтогоВсеЛитры;
	Для Каждого КодТ Из КодыТоплива Цикл
		БезналЛитры = БезналЛитры - РозницаПоТопливу.Получить(КодТ).Литры;
	КонецЦикла;

	// ===== ГЕНЕРАЦИЯ HTML =====

	HTML = СтильСмены();

	// --- ШАПКА ---
	// Главная цифра — учётная выручка (= розница). Безнал показываем литрами «к расценке».
	HTML = HTML + "<div class=""header""><div>"
		+ "<div class=""header-station"">" + Экр(Заголовок) + "</div>"
		+ "</div><div>"
		+ "<div class=""header-total"">" + Формат(РозницаИтого, "ЧДЦ=2; ЧРД=,; ЧГ=' '") + " &#8381;</div>"
		+ "<div class=""header-sub"">учётная выручка &middot; "
		+ Формат(ИтогоВсеЛитры, "ЧДЦ=2") + " л &middot; "
		+ XMLСтрока(КолТоплива) + " вида топлива</div>";
	Если БезналЛитры > 0 Тогда
		HTML = HTML + "<div class=""header-sub"" style=""color:#888"">"
			+ Формат(БезналЛитры, "ЧДЦ=2") + " л безналично — к расценке в конце месяца</div>";
	КонецЕсли;
	HTML = HTML + "</div></div>";

	// --- ТАБЛИЦА ПРОДАЖ ---
	HTML = HTML + "<div class=""section"">Сменный отчёт &mdash; продажи по видам оплаты</div>"
		+ "<table><tr><th rowspan=""2"">Вид оплаты</th><th rowspan=""2"">Маршрут</th>"
		+ "<th rowspan=""2"">Проводка</th>";

	Для Каждого КодТ Из КодыТоплива Цикл
		HTML = HTML + "<th colspan=""2"">" + Экр(КодТ) + "</th>";
	КонецЦикла;
	HTML = HTML + "<th rowspan=""2"">Итого</th></tr><tr>";
	Для Каждого КодТ Из КодыТоплива Цикл
		HTML = HTML + "<th>Литры</th><th>Сумма</th>";
	КонецЦикла;
	HTML = HTML + "</tr>";

	ПеремещенияHTML = "";

	Для Каждого КЗ Из АгрегатОплат Цикл
		ИмяОпл = КЗ.Ключ;
		ДанныеОпл = КЗ.Значение;
		ИмяНРег = НРег(ИмяОпл);
		ЭтоРозница = КешЭтоРозница.Получить(ИмяОпл);

		// Определить маршрут и проводку
		Если СтрНайти(ИмяНРег, "наличн") > 0 Тогда
			Маршрут = "<span class=""route-tag route-retail"">Розница</span>";
			Проводка = "Дт 50.01 Кт 62.Р";
		ИначеЕсли СтрНайти(ИмяНРег, "сбербанк") > 0 Тогда
			Маршрут = "<span class=""route-tag route-retail"">Розница</span>";
			Проводка = "Дт 57.03 Кт 62.Р";
		ИначеЕсли СтрНайти(ИмяНРег, "мобил") > 0 ИЛИ СтрНайти(ИмяНРег, "яндекс") > 0
			ИЛИ СтрНайти(ИмяНРег, "benzuber") > 0 Тогда
			Маршрут = "<span class=""route-tag route-transfer"">ЯНДЕКС</span>";
			Проводка = "Перемещение → ЯНДЕКС";
		ИначеЕсли СтрНайти(ИмяНРег, "ведомост") > 0 Тогда
			Маршрут = "<span class=""route-tag route-transfer"">Ведомости</span>";
			Проводка = "Перемещение → Ведомости";
		ИначеЕсли СтрНайти(ИмяНРег, "талон") > 0 Тогда
			Маршрут = "<span class=""route-tag route-transfer"">Талоны</span>";
			Проводка = "Перемещение → Талоны";
		ИначеЕсли СтрНайти(ИмяНРег, "корп") > 0 ИЛИ СтрНайти(ИмяНРег, "агора") > 0
			ИЛИ СтрНайти(ИмяНРег, "топливн") > 0
			ИЛИ СтрНайти(ИмяНРег, "viacard") > 0
			ИЛИ СтрНайти(ИмяНРег, "балтоп") > 0
			ИЛИ СтрНайти(ИмяНРег, "baltop") > 0
			ИЛИ ИмяНРег = "кр" Тогда
			Маршрут = "<span class=""route-tag route-transfer"">Карты</span>";
			Проводка = "Перемещение → Карты";
		Иначе
			Маршрут = Экр(ИмяОпл);
			Проводка = "";
		КонецЕсли;

		СтрокаHTML = "<tr><td>" + Экр(ИмяОпл) + "</td><td>" + Маршрут + "</td>"
			+ "<td class=""provodka"">" + Проводка + "</td>";

		ИтогоСтроки = 0;
		ЕстьЛитрыВСтроке = Ложь;
		Для Каждого КодТ Из КодыТоплива Цикл
			Яч = ДанныеОпл.Получить(КодТ);
			Если Яч <> Неопределено И Яч.Литры > 0 Тогда
				ЕстьЛитрыВСтроке = Истина;
				СтрокаHTML = СтрокаHTML + "<td class=""num"">" + Формат(Яч.Литры, "ЧДЦ=2") + "</td>";
				Если ЭтоРозница Тогда
					СтрокаHTML = СтрокаHTML + "<td class=""num"">" + Формат(Яч.Сумма, "ЧДЦ=2") + "</td>";
					ИтогоСтроки = ИтогоСтроки + Яч.Сумма;
				Иначе
					СтрокаHTML = СтрокаHTML + "<td class=""num"" style=""color:#aaa"">&mdash;</td>";
				КонецЕсли;
			Иначе
				СтрокаHTML = СтрокаHTML + "<td class=""num"">&mdash;</td><td class=""num"">&mdash;</td>";
			КонецЕсли;
		КонецЦикла;
		Если ЭтоРозница Тогда
			СтрокаHTML = СтрокаHTML + "<td class=""num""><b>" + Формат(ИтогоСтроки, "ЧДЦ=2") + "</b></td></tr>";
		Иначе
			СтрокаHTML = СтрокаHTML + "<td class=""num"" style=""color:#aaa"">&mdash;</td></tr>";
		КонецЕсли;

		Если ЭтоРозница Тогда
			HTML = HTML + СтрокаHTML;
		ИначеЕсли ЕстьЛитрыВСтроке Тогда
			ПеремещенияHTML = ПеремещенияHTML + СтрокаHTML;
		КонецЕсли;
	КонецЦикла;

	// Строка итого розницы
	HTML = HTML + "<tr class=""total-row""><td colspan=""3"">ИТОГО розница</td>";
	Для Каждого КодТ Из КодыТоплива Цикл
		Р = РозницаПоТопливу.Получить(КодТ);
		HTML = HTML + "<td class=""num"">" + Формат(Р.Литры, "ЧДЦ=2") + "</td>"
			+ "<td class=""num"">" + Формат(Р.Сумма, "ЧДЦ=2") + "</td>";
	КонецЦикла;
	HTML = HTML + "<td class=""num""><b>" + Формат(РозницаИтого, "ЧДЦ=2") + "</b></td></tr>";

	// Перемещения
	Если ЗначениеЗаполнено(ПеремещенияHTML) Тогда
		HTML = HTML + ПеремещенияHTML;
	КонецЕсли;

	// Строка итого за смену — только литры. Сумма по смене смешана быть не может
	// (часть розничная учётная, часть безналичная без цены — суммировать нельзя).
	HTML = HTML + "<tr class=""total-row""><td colspan=""3"">ИТОГО за смену (литры)</td>";
	Для Каждого КодТ Из КодыТоплива Цикл
		Итог = ИтогиПоТопливу.Получить(КодТ);
		HTML = HTML + "<td class=""num""><b>" + Формат(Итог.Литры, "ЧДЦ=2") + "</b></td>"
			+ "<td class=""num"" style=""color:#aaa"">&mdash;</td>";
	КонецЦикла;
	HTML = HTML + "<td class=""num"" style=""color:#aaa"">&mdash;</td></tr></table>";

	// --- ПРОВОДКИ ---
	HTML = HTML + "<table style=""margin-bottom:2px"">"
		+ "<tr><th colspan=""3"" style=""text-align:left"">Проводки при проведении</th></tr>"
		+ "<tr><td class=""provodka"" style=""width:220px"">Дт 62.Р Кт 90.01.1</td>"
		+ "<td>Выручка (по каждому виду топлива)</td>"
		+ "<td class=""num"">" + Формат(РозницаИтого, "ЧДЦ=2") + "</td></tr>"
		+ "<tr><td class=""provodka"">Дт 90.02.1 Кт 41.02</td>"
		+ "<td>Списание себестоимости с розничного склада</td>"
		+ "<td class=""num"">по учётной цене</td></tr>"
		+ "<tr><td class=""provodka"">Дт 50.01 Кт 62.Р</td>"
		+ "<td>Оплата наличными</td><td class=""num""></td></tr>"
		+ "<tr><td class=""provodka"">Дт 57.03 Кт 62.Р</td>"
		+ "<td>Оплата картой (эквайринг)</td><td class=""num""></td></tr>"
		+ "<tr><td class=""provodka"">Дт 90.03 Кт 68.02</td>"
		+ "<td>НДС 22%</td>"
		+ "<td class=""num"">" + Формат(Окр(РозницаИтого * 22 / 122, 2), "ЧДЦ=2") + "</td></tr>"
		+ "</table>";

	// --- РЕЗЕРВУАРЫ ---
	МассивРезервуаров = TL_ApiКлиент.ПолучитьЗначение(Данные, "release", Неопределено);
	Если ТипЗнч(МассивРезервуаров) = Тип("Массив") И МассивРезервуаров.Количество() > 0 Тогда
		HTML = HTML + "<div class=""section"" style=""margin-top:6px"">Резервуары</div><div class=""tanks-row"">";

		Для Каждого Рез Из МассивРезервуаров Цикл
			НомерБака = XMLСтрока(TL_ApiКлиент.ПолучитьЗначение(Рез, "tank", 0));
			СервисР = TL_ApiКлиент.ПолучитьЗначение(Рез, "service", Неопределено);
			ИмяТоплР = "";
			Если СервисР <> Неопределено Тогда
				ИмяТоплР = Строка(TL_ApiКлиент.ПолучитьЗначение(СервисР, "service_name", ""));
			КонецЕсли;

			ДокНач = TL_ApiКлиент.ПолучитьЗначение(Рез, "doc_beg", Неопределено);
			НачалоЛ = ?(ДокНач <> Неопределено, Число(TL_ApiКлиент.ПолучитьЗначение(ДокНач, "volume", 0)), 0);

			ОстатокР = TL_ApiКлиент.ПолучитьЗначение(Рез, "rest", Неопределено);
			ОстатокЛ = ?(ОстатокР <> Неопределено, Число(TL_ApiКлиент.ПолучитьЗначение(ОстатокР, "volume", 0)), 0);

			РасходР = TL_ApiКлиент.ПолучитьЗначение(Рез, "release", Неопределено);
			РасходЛ = ?(РасходР <> Неопределено, Число(TL_ApiКлиент.ПолучитьЗначение(РасходР, "volume", 0)), 0);

			ПлотнКон = Число(TL_ApiКлиент.ПолучитьЗначение(Рез, "density_end", 0));
			ТемпКон = Число(TL_ApiКлиент.ПолучитьЗначение(Рез, "temp_end", 0));

			// CSS-класс цвета
			ИмяНРег = НРег(ИмяТоплР);
			Если СтрНайти(ИмяНРег, "92") > 0 Тогда
				КлассЦвета = "fill-92"; КлассИмени = "fuel-92";
			ИначеЕсли СтрНайти(ИмяНРег, "95") > 0 Тогда
				КлассЦвета = "fill-95"; КлассИмени = "fuel-95";
			Иначе
				КлассЦвета = "fill-dt"; КлассИмени = "fuel-dt";
			КонецЕсли;

			// Процент заполнения (условно макс = начало × 1.5)
			Макс = ?(НачалоЛ > 0, НачалоЛ * 1.5, 10000);
			ПроцНач = Мин(Окр(НачалоЛ / Макс * 100), 100);
			ПроцОст = Мин(Окр(ОстатокЛ / Макс * 100), 100);

			HTML = HTML + "<div class=""tank"">"
				+ "<div class=""tank-head""><span>Бак " + НомерБака + "</span>"
				+ " <span class=""" + КлассИмени + """>" + Экр(ИмяТоплР) + "</span></div>"
				+ "<div class=""tank-bar"">"
				+ "<div class=""tank-bar-used " + КлассЦвета + """ style=""width:" + XMLСтрока(ПроцНач) + "%""></div>"
				+ "<div class=""tank-bar-remain " + КлассЦвета + """ style=""width:" + XMLСтрока(ПроцОст) + "%""></div>"
				+ "<span class=""tank-bar-label lbl-left"" style=""color:#fff"">" + Формат(ОстатокЛ, "ЧДЦ=2") + " л</span>"
				+ "<span class=""tank-bar-label lbl-right"">" + Формат(НачалоЛ, "ЧДЦ=2") + " л</span>"
				+ "</div>"
				+ "<table class=""tank-stats""><tr>"
				+ "<td>Отпущено</td><td class=""num negative"">&minus;" + Формат(РасходЛ, "ЧДЦ=2") + " л</td>"
				+ "<td>&rho;</td><td class=""num"">" + ?(ПлотнКон > 0, Формат(ПлотнКон, "ЧДЦ=4"), "-") + "</td>"
				+ "<td>t</td><td class=""num"">" + ?(ТемпКон > 0, Формат(ТемпКон, "ЧДЦ=1") + "&deg;C", "-") + "</td>"
				+ "</tr></table></div>";
		КонецЦикла;

		HTML = HTML + "</div>";
	КонецЕсли;

	// --- ДОКУМЕНТЫ 1С ---
	HTML = HTML + "<div class=""section"">Документы 1С</div><div>";
	Если СвязанныеДокументы <> Неопределено И СвязанныеДокументы.Количество() > 0 Тогда
		Для Каждого Док Из СвязанныеДокументы Цикл
			Класс = ?(Док.Статус = "Проведён", "doc-row checked", "doc-row");
			ССылкаОткр = "";
			Если Док.Свойство("UUID") И ЗначениеЗаполнено(Док.UUID)
					И Док.Свойство("ВидДокумента") И ЗначениеЗаполнено(Док.ВидДокумента) Тогда
				ССылкаОткр = " <a href=""#"" data-action=""open|" + Док.ВидДокумента + "|" + Док.UUID
					+ """ style=""margin-left:8px;font-size:11px;text-decoration:none;""> открыть</a>";
			КонецЕсли;
			Маршрут = "";
			Если Док.Свойство("Маршрут") И ЗначениеЗаполнено(Док.Маршрут) Тогда
				Маршрут = " <span style=""color:#6B7280;font-size:11px;"">" + Экр(Док.Маршрут) + "</span>";
			КонецЕсли;
			HTML = HTML + "<div class=""" + Класс + """>"
				+ "<div class=""doc-check"">&#10003;</div>"
				+ "<div class=""doc-name""><b>" + Экр(Док.ТипДокумента) + "</b>" + Маршрут
				+ " <span style=""color:#1F2937;"">" + Экр(Док.Описание) + "</span>"
				+ ССылкаОткр + "</div>"
				+ "</div>";
		КонецЦикла;
	Иначе
		HTML = HTML + "<div class=""doc-row disabled""><div class=""doc-name"">Ещё не загружено в 1С</div></div>";
	КонецЕсли;
	HTML = HTML + "</div>";

	HTML = HTML + "</body></html>";

	Возврат HTML;

КонецФункции

// Сформировать HTML деталей для ТТН (цепочка документов).
//
// Параметры:
//   ДанныеJSON          - Строка - JSON поступления
//   Заголовок           - Строка
//   СвязанныеДокументы  - Массив из Структура("ТипДокумента, Описание, Статус")
//
Функция СформироватьДеталиТТН(ДанныеJSON, Заголовок, СвязанныеДокументы = Неопределено) Экспорт

	Попытка
		Данные = TL_СозданиеДокументов.ПолучитьИзJSON(ДанныеJSON);
	Исключение
		Возврат ОбёрткаHTML("<p style=""color:#c00;"">Ошибка разбора JSON</p>");
	КонецПопытки;

	Если Данные = Неопределено Тогда
		Возврат ОбёрткаHTML("<p style=""color:#888;"">Нет данных</p>");
	КонецЕсли;

	КодТоплива = Строка(TL_ApiКлиент.ПолучитьЗначение(Данные, "fuel", ""));
	ИмяТоплива = TL_Настройки.ПолучитьНаименованиеТоплива(КодТоплива);
	Плотность = Число(TL_ApiКлиент.ПолучитьЗначение(Данные, "density", 0));
	НомерТТН = Строка(TL_ApiКлиент.ПолучитьЗначение(Данные, "ttn", ""));

	ДокДанные = TL_ApiКлиент.ПолучитьЗначение(Данные, "doc", Неопределено);
	Если ДокДанные <> Неопределено Тогда
		Литры = Число(TL_ApiКлиент.ПолучитьЗначение(ДокДанные, "volume", 0));
		МассаКг = Число(TL_ApiКлиент.ПолучитьЗначение(ДокДанные, "amount", 0));
		Стоимость = Число(TL_ApiКлиент.ПолучитьЗначение(ДокДанные, "cost", 0));
	Иначе
		Литры = Число(TL_ApiКлиент.ПолучитьЗначение(Данные, "doc_volume", 0));
		МассаКг = Число(TL_ApiКлиент.ПолучитьЗначение(Данные, "doc_amount", 0));
		Стоимость = Число(TL_ApiКлиент.ПолучитьЗначение(Данные, "amount", 0));
	КонецЕсли;
	МассаТонн = МассаКг / 1000;

	HTML = СтильСмены();

	// --- ШАПКА ---
	HTML = HTML + "<div class=""header""><div>"
		+ "<div class=""header-station"">" + Экр(Заголовок) + "</div>"
		+ "</div><div>"
		+ "<div class=""header-total"">" + Формат(МассаТонн, "ЧДЦ=3") + " т</div>"
		+ "<div class=""header-sub"">" + Экр(ИмяТоплива) + " &middot; &rho;="
		+ ?(Плотность > 0, Формат(Плотность, "ЧДЦ=4"), "-") + "</div>"
		+ "</div></div>";

	// --- ДАННЫЕ ТТН ---
	HTML = HTML + "<div class=""section"">Данные ТТН</div>"
		+ "<table><tr><th>Параметр</th><th>Значение</th></tr>"
		+ "<tr><td>Топливо</td><td><b>" + Экр(ИмяТоплива) + "</b></td></tr>"
		+ "<tr><td>Масса</td><td class=""num""><b>" + Формат(МассаТонн, "ЧДЦ=3") + " т</b> (" + Формат(МассаКг, "ЧДЦ=2") + " кг)</td></tr>"
		+ "<tr><td>Объём (книжный)</td><td class=""num""><b>" + Формат(Литры, "ЧДЦ=2") + " л</b></td></tr>"
		+ "<tr><td>Плотность</td><td class=""num"">" + ?(Плотность > 0, Формат(Плотность, "ЧДЦ=4") + " кг/л", "&mdash;") + "</td></tr>"
		+ "</table>";

	// --- ЦЕПОЧКА ДОКУМЕНТОВ ---
	// Если переданы СвязанныеДокументы — извлекаем UUID для кнопки « открыть»
	UUIDПерем = "";
	UUIDКомпл = "";
	Если ТипЗнч(СвязанныеДокументы) = Тип("Массив") Тогда
		Для Каждого СвязДок Из СвязанныеДокументы Цикл
			Если СвязДок.Свойство("UUID") И ЗначениеЗаполнено(СвязДок.UUID) Тогда
				Если ВРег(СвязДок.ТипДокумента) = "ПЕРЕМЕЩЕНИЕ" Тогда
					UUIDПерем = СвязДок.UUID;
				ИначеЕсли ВРег(СвязДок.ТипДокумента) = "КОМПЛЕКТАЦИЯ" Тогда
					UUIDКомпл = СвязДок.UUID;
				КонецЕсли;
			КонецЕсли;
		КонецЦикла;
	КонецЕсли;

	HTML = HTML + "<div class=""section"">Документы для создания</div><div>";

	// Шаг 1: Перемещение
	HTML = HTML + "<div class=""doc-row checked"">"
		+ "<div class=""doc-check"">&#10003;</div>"
		+ "<div class=""doc-name""><b>Перемещение</b> Осн.склад &rarr; АЗС"
		+ ?(ЗначениеЗаполнено(UUIDПерем),
			" <a href=""#"" data-action=""open|Перемещение|" + UUIDПерем
				+ """ style=""margin-left:8px;font-size:11px;text-decoration:none;""> открыть</a>",
			"")
		+ "</div>"
		+ "<div class=""doc-detail"">Дт 41.01(АЗС) Кт 41.01(Основной) &middot; "
		+ Экр(ИмяТоплива) + " (т) " + Формат(МассаТонн, "ЧДЦ=3") + " т</div>"
		+ "<div class=""doc-amount"">" + Формат(МассаТонн, "ЧДЦ=3") + " т</div></div>";

	// Шаг 2: Комплектация
	HTML = HTML + "<div class=""doc-row checked"">"
		+ "<div class=""doc-check"">&#10003;</div>"
		+ "<div class=""doc-name""><b>Комплектация</b> тонны &rarr; литры"
		+ ?(ЗначениеЗаполнено(UUIDКомпл),
			" <a href=""#"" data-action=""open|Комплектация|" + UUIDКомпл
				+ """ style=""margin-left:8px;font-size:11px;text-decoration:none;""> открыть</a>",
			"")
		+ "</div>"
		+ "<div class=""doc-detail"">Кт 41.01(тонны) &rarr; Дт 41.02(литры) &middot; " + Экр(ИмяТоплива) + " (т) "
		+ Формат(МассаТонн, "ЧДЦ=3") + " т &rarr; " + Экр(ИмяТоплива) + " (л) "
		+ Формат(Литры, "ЧДЦ=2") + " л &middot; &rho;="
		+ ?(Плотность > 0, Формат(Плотность, "ЧДЦ=4"), "-") + "</div>"
		+ "<div class=""doc-amount"">" + Формат(Литры, "ЧДЦ=2") + " л</div></div>";

	HTML = HTML + "</div></body></html>";

	Возврат HTML;

КонецФункции

// Сформировать HTML-страницу настроек (замена мастера TL_НастройкаРасширения).
// Все настройки редактируются inline. Кнопка "Сохранить" генерирует click-событие,
// данные передаются через window.settingsJSON.
//
// Параметры:
//   ТекущиеНастройки - Строка - JSON с текущими значениями из регистра TL_Настройки
//
// Возвращаемое значение:
//   Строка - HTML документ
//
Функция СформироватьНастройки(ТекущиеНастройки = "") Экспорт

	HTML = "<!DOCTYPE html><html><head><meta charset=""utf-8"">
	|<style>
	|* { margin:0; padding:0; box-sizing:border-box; }
	|body { font-family:-apple-system,'Segoe UI',sans-serif; background:#F0F5FA; padding:16px; font-size:13px; }
	|h2 { color:#1F2937; font-size:16px; margin-bottom:12px; }
	|h3 { color:#3B82F6; font-size:13px; text-transform:uppercase; letter-spacing:1px; margin:0 0 10px; }
	|.card { background:#fff; border-radius:12px; padding:16px; margin-bottom:12px;
	|  box-shadow:0 1px 3px rgba(0,0,0,0.08); }
	|.grid2 { display:grid; grid-template-columns:1fr 1fr; gap:8px 16px; }
	|.row { display:flex; gap:8px; margin-bottom:6px; align-items:center; }
	|.row label { min-width:120px; color:#6B7280; font-size:12px; }
	|.row input, .row select { flex:1; padding:5px 8px; border:1px solid #D9DFE7;
	|  border-radius:6px; font-size:13px; }
	|.row input:focus { outline:none; border-color:#3B82F6; box-shadow:0 0 0 2px rgba(59,130,246,0.15); }
	|table { width:100%; border-collapse:collapse; }
	|th { text-align:left; font-size:11px; color:#6B7280; text-transform:uppercase;
	|  padding:4px 6px; border-bottom:2px solid #D9DFE7; }
	|td { padding:4px 6px; border-bottom:1px solid #E5E7EB; }
	|td input, td select { width:100%; padding:3px 6px; border:1px solid #D9DFE7; border-radius:4px; font-size:12px; }
	|.badge { display:inline-block; padding:1px 6px; border-radius:10px; font-size:11px; font-weight:500; }
	|.badge-ok { background:#DCFCE7; color:#16A34A; }
	|.badge-warn { background:#FEF3C7; color:#D97706; }
	|.badge-fuel { background:#EFF6FF; color:#3B82F6; margin:1px 2px; }
	|.btn { padding:7px 20px; border:none; border-radius:8px; font-size:13px;
	|  font-weight:600; cursor:pointer; }
	|.btn-primary { background:#3B82F6; color:#fff; }
	|.btn-primary:hover { background:#2563EB; }
	|.btn-sm { padding:4px 12px; font-size:12px; border-radius:6px; }
	|.btn-secondary { background:#E5E7EB; color:#374151; }
	|.sep { border-top:1px solid #E5E7EB; margin:12px 0; }
	|</style></head><body>
	|<h2>Настройки ElsyPlus Ledger</h2>
	|
	|<div class=""card"">
	|<h3>Подключение</h3>
	|<div class=""grid2"">
	|<div class=""row""><label>URL API</label><input id=""s_url"" value=""""></div>
	|<div class=""row""><label>Код системы</label><input id=""s_system"" value="""" style=""max-width:80px;""></div>
	|<div class=""row""><label>Логин</label><input id=""s_login"" value=""""></div>
	|<div class=""row""><label>Пароль</label><input id=""s_password"" type=""password"" value=""""></div>
	|</div>
	|<div class=""sep""></div>
	|<div class=""grid2"">
	|<div class=""row""><label>Организация</label><input id=""s_org"" value=""""></div>
	|<div class=""row""><label>Основной склад</label><input id=""s_warehouse"" value=""""></div>
	|</div>
	|</div>
	|
	|<div class=""card"">
	|<h3>Станции</h3>
	|<table>
	|<tr><th style=""width:55px"">Код</th><th>Наименование</th><th>Склад 1С</th><th style=""width:75px"">Закр.</th><th>Виды топлива</th></tr>
	|<tbody id=""tblStations""></tbody>
	|</table>
	|<button class=""btn btn-secondary btn-sm"" onclick=""addStation()"" style=""margin-top:8px;"">+ Станция</button>
	|</div>
	|
	|<div class=""card"">
	|<h3>Каналы оплат</h3>
	|<table>
	|<tr><th style=""width:70px"">Канал</th><th>Наименование</th><th>Склад</th><th style=""width:80px"">Перемещ.</th></tr>
	|<tbody id=""tblPayments""></tbody>
	|</table>
	|<div class=""sep""></div>
	|<h3>Номенклатура топлива</h3>
	|<table>
	|<tr><th style=""width:55px"">Код</th><th>Название</th><th>Номенклатура (т)</th><th>Номенклатура (л)</th><th style=""width:80px"">Плотность</th></tr>
	|<tbody id=""tblFuel""></tbody>
	|</table>
	|<button class=""btn btn-secondary btn-sm"" onclick=""addFuel()"" style=""margin-top:8px;"">+ Топливо</button>
	|</div>
	|
	|<div style=""display:flex; gap:12px; margin-top:12px;"">
	|<button class=""btn btn-primary"" id=""btnSave"">Сохранить</button>
	|</div>
	|
	|<script>
	|var cfg = {};
	|try { cfg = JSON.parse('" + Экр(ТекущиеНастройки) + "'); } catch(e) { cfg = {}; }
	|function v(key, def) { return cfg[key] || def || ''; }
	|var fuelNames = {'2':'АИ-92','3':'АИ-95','5':'ДТ','100':'АИ-100','98':'АИ-98'};
	|
	|// Подключение
	|document.getElementById('s_url').value = v('URLСервера');
	|document.getElementById('s_login').value = v('Логин');
	|document.getElementById('s_password').value = v('Пароль');
	|document.getElementById('s_system').value = v('КодСистемы');
	|document.getElementById('s_org').value = v('Организация');
	|document.getElementById('s_warehouse').value = v('ОсновнойСклад');
	|
	|// Собрать виды топлива (для колонки в таблице станций)
	|var fuelCodes = [];
	|Object.keys(cfg).forEach(function(k) {
	|  var m = k.match(/^Топливо_(\w+)_Тонны$/);
	|  if (m) fuelCodes.push(m[1]);
	|});
	|function fuelBadges() {
	|  return fuelCodes.map(function(c) {
	|    return '<span class=""badge badge-fuel"">' + (fuelNames[c] || c) + '</span>';
	|  }).join('');
	|}
	|
	|// Станции — динамический скан
	|var stations = [];
	|Object.keys(cfg).forEach(function(k) {
	|  var m = k.match(/^Станция_(\d+)_Наименование$/);
	|  if (m) stations.push(m[1]);
	|});
	|if (stations.length === 0) stations = ['5'];
	|stations.sort(function(a,b) { return Number(a)-Number(b); });
	|stations.forEach(function(code) { addStation(code); });
	|
	|function addStation(code) {
	|  code = code || '';
	|  var tb = document.getElementById('tblStations');
	|  var tr = document.createElement('tr');
	|  var wh = v('Станция_'+code+'_Склад');
	|  var whStyle = wh ? '' : 'background:#FEF2F2;';
	|  var whBadge = wh ? '<span class=""badge badge-ok"">OK</span>' : '<span class=""badge badge-warn"">нет</span>';
	|  tr.innerHTML = '<td><input class=""st_code"" value=""'+code+'"" style=""width:50px;text-align:center""></td>'
	|    + '<td><input class=""st_name"" value=""'+v('Станция_'+code+'_Наименование')+'""></td>'
	|    + '<td style=""'+whStyle+'""><input class=""st_wh"" value=""'+wh+'"" placeholder=""Имя склада в 1С""></td>'
	|    + '<td><input type=""time"" class=""st_close"" value=""'+v('Станция_'+code+'_ВремяЗакрытия','00:00')+'""></td>'
	|    + '<td style=""white-space:nowrap"">'+fuelBadges()+'</td>';
	|  tb.appendChild(tr);
	|}
	|
	|// Топливо
	|if (fuelCodes.length === 0) fuelCodes = ['2','3','5'];
	|fuelCodes.forEach(function(code) { addFuel(code); });
	|
	|function addFuel(code) {
	|  code = code || '';
	|  var tb = document.getElementById('tblFuel');
	|  var tr = document.createElement('tr');
	|  tr.innerHTML = '<td><input class=""f_code"" value=""'+code+'"" style=""width:50px;text-align:center""></td>'
	|    + '<td><input class=""f_name"" value=""'+(fuelNames[code]||'')+'"" readonly style=""background:#F9FAFB;color:#6B7280""></td>'
	|    + '<td><input class=""f_tons"" value=""'+v('Топливо_'+code+'_Тонны')+'""></td>'
	|    + '<td><input class=""f_litres"" value=""'+v('Топливо_'+code+'_Литры')+'""></td>'
	|    + '<td><input class=""f_density"" value=""'+v('Топливо_'+code+'_Плотность')+'"" style=""width:75px;text-align:center""></td>';
	|  tb.appendChild(tr);
	|}
	|
	|// Каналы оплат
	|var payments = [
	|  {key:'retail', name:v('Оплата_retail_Наименование','Розница (наличные + эквайринг)')},
	|  {key:'cards', name:v('Оплата_cards_Наименование','Топливные/корп. карты')},
	|  {key:'online', name:v('Оплата_online_Наименование','Онлайн (Яндекс, МобилПр.)')},
	|  {key:'ledger', name:v('Оплата_ledger_Наименование','Ведомости')}
	|];
	|var ptb = document.getElementById('tblPayments');
	|payments.forEach(function(p) {
	|  var tr = document.createElement('tr');
	|  var req = v('Оплата_'+p.key+'_ТребуетПеремещения','Нет');
	|  tr.innerHTML = '<td style=""font-weight:600;color:#374151"">'+p.key+'</td>'
	|    + '<td><input class=""p_name"" data-key=""'+p.key+'"" value=""'+p.name+'""></td>'
	|    + '<td><input class=""p_wh"" data-key=""'+p.key+'"" value=""'+v('Оплата_'+p.key+'_Склад')+'"" placeholder=""—""></td>'
	|    + '<td><select class=""p_move"" data-key=""'+p.key+'""><option'+(req==='Да'?' selected':'')+'>Да</option><option'+(req!=='Да'?' selected':'')+'>Нет</option></select></td>';
	|  ptb.appendChild(tr);
	|});
	|
	|// Сохранение
	|document.getElementById('btnSave').addEventListener('click', function() {
	|  var result = {};
	|  result['URLСервера'] = document.getElementById('s_url').value;
	|  result['Логин'] = document.getElementById('s_login').value;
	|  result['Пароль'] = document.getElementById('s_password').value;
	|  result['КодСистемы'] = document.getElementById('s_system').value;
	|  result['Организация'] = document.getElementById('s_org').value;
	|  result['ОсновнойСклад'] = document.getElementById('s_warehouse').value;
	|
	|  var rows = document.getElementById('tblStations').rows;
	|  for (var i = 0; i < rows.length; i++) {
	|    var code = rows[i].querySelector('.st_code').value;
	|    if (!code) continue;
	|    result['Станция_'+code+'_Наименование'] = rows[i].querySelector('.st_name').value;
	|    result['Станция_'+code+'_Склад'] = rows[i].querySelector('.st_wh').value;
	|    result['Станция_'+code+'_ВремяЗакрытия'] = rows[i].querySelector('.st_close').value;
	|  }
	|
	|  rows = document.getElementById('tblFuel').rows;
	|  for (var i = 0; i < rows.length; i++) {
	|    var code = rows[i].querySelector('.f_code').value;
	|    if (!code) continue;
	|    result['Топливо_'+code+'_Тонны'] = rows[i].querySelector('.f_tons').value;
	|    result['Топливо_'+code+'_Литры'] = rows[i].querySelector('.f_litres').value;
	|    result['Топливо_'+code+'_Плотность'] = rows[i].querySelector('.f_density').value;
	|  }
	|
	|  document.querySelectorAll('.p_name').forEach(function(el) {
	|    result['Оплата_'+el.dataset.key+'_Наименование'] = el.value;
	|  });
	|  document.querySelectorAll('.p_wh').forEach(function(el) {
	|    result['Оплата_'+el.dataset.key+'_Склад'] = el.value;
	|  });
	|  document.querySelectorAll('.p_move').forEach(function(el) {
	|    result['Оплата_'+el.dataset.key+'_ТребуетПеремещения'] = el.value;
	|  });
	|
	|  window.settingsJSON = JSON.stringify(result);
	|  setTimeout(function() { document.getElementById('btnSave').click(); }, 1);
	|});
	|</script>
	|</body></html>";

	Возврат СтрЗаменить(HTML, Символы.ПС + "|", Символы.ПС);

КонецФункции

// Пустой HTML (при отсутствии выбора).
//
Функция ПустыеДетали() Экспорт
	Возврат ОбёрткаHTML(
		"<div style=""text-align:center; padding:40px; color:#D1D5DB;"">"
		+ "<div style=""font-size:36px;""></div>"
		+ "<div style=""margin-top:8px;"">Выберите строку для просмотра деталей</div>"
		+ "</div>"
	);
КонецФункции

// Сформировать детальный HTML-отчёт ошибок загрузки с группировкой по категориям
// и рекомендациями «что делать» для бухгалтера. F-направление плана.
//
// Параметры:
//   ИдентификаторПакета - Строка - ключ пакета/смены. Если пусто — показывает все ошибки за сегодня.
//   Заголовок          - Строка - заголовок HTML-страницы
//
// Возвращаемое значение:
//   Строка - готовый HTML
//
Функция СформироватьДетальныйОтчётОшибок(ИдентификаторПакета = "", Заголовок = "Журнал ошибок загрузки", МоментОт = Неопределено) Экспорт

	// 1. Сбор ошибок из TL_ОшибкиЗагрузки
	Запрос = Новый Запрос;
	Если ЗначениеЗаполнено(ИдентификаторПакета) Тогда
		Запрос.Текст =
			"ВЫБРАТЬ
			|	Т.ИдентификаторПакета, Т.НомерЗаписи, Т.ИсточникUUID,
			|	Т.ТипОбъекта, Т.Поле, Т.КодОшибки,
			|	Т.ЗначениеИсточника, Т.СообщениеОшибки, Т.ВремяРегистрации, Т.Документ
			|ИЗ РегистрСведений.TL_ОшибкиЗагрузки КАК Т
			|ГДЕ Т.ИдентификаторПакета = &Пакет
			|УПОРЯДОЧИТЬ ПО Т.ВремяРегистрации, Т.НомерЗаписи";
		Запрос.УстановитьПараметр("Пакет", ИдентификаторПакета);
	Иначе
		// МоментОт — если передан, фильтруем от него (для «текущая сессия» по
		// кнопке Обновить). Иначе — за весь сегодня.
		НачальныйМомент = ?(ЗначениеЗаполнено(МоментОт), МоментОт, НачалоДня(ТекущаяДата()));
		Запрос.Текст =
			"ВЫБРАТЬ ПЕРВЫЕ 200
			|	Т.ИдентификаторПакета, Т.НомерЗаписи, Т.ИсточникUUID,
			|	Т.ТипОбъекта, Т.Поле, Т.КодОшибки,
			|	Т.ЗначениеИсточника, Т.СообщениеОшибки, Т.ВремяРегистрации, Т.Документ
			|ИЗ РегистрСведений.TL_ОшибкиЗагрузки КАК Т
			|ГДЕ Т.ВремяРегистрации >= &МоментОт
			|УПОРЯДОЧИТЬ ПО Т.ВремяРегистрации УБЫВ, Т.НомерЗаписи";
		Запрос.УстановитьПараметр("МоментОт", НачальныйМомент);
	КонецЕсли;

	ПоКатегориям = Новый Соответствие; // КодОшибки → Массив(Структура)
	Попытка
		Выборка = Запрос.Выполнить().Выбрать();
	Исключение
		ТекстОшибки = ОписаниеОшибки();
		Возврат ОбёрткаHTML("<div style=""padding:20px;color:#888"">"
			+ "<b>Регистр TL_ОшибкиЗагрузки недоступен.</b><br>"
			+ "<pre style=""color:#b91c1c;font-size:11px;margin-top:8px;white-space:pre-wrap;"">"
			+ ТекстОшибки + "</pre></div>");
	КонецПопытки;

	ВсегоОшибок = 0;
	Пока Выборка.Следующий() Цикл
		ВсегоОшибок = ВсегоОшибок + 1;
		Код = Выборка.КодОшибки;
		Список = ПоКатегориям.Получить(Код);
		Если Список = Неопределено Тогда
			Список = Новый Массив;
			ПоКатегориям.Вставить(Код, Список);
		КонецЕсли;
		Список.Добавить(Новый Структура(
			"Пакет, Поле, Значение, Сообщение, Время, Документ",
			Выборка.ИдентификаторПакета, Выборка.Поле, Выборка.ЗначениеИсточника,
			Выборка.СообщениеОшибки, Выборка.ВремяРегистрации, Выборка.Документ));
	КонецЦикла;

	// 2. Формирование HTML
	HTML = СтильОшибок();
	HTML = HTML + "<div class=""err-page"">";
	HTML = HTML + "<h1>" + Экр(Заголовок) + "</h1>";
	HTML = HTML + "<div class=""err-summary"">Всего ошибок: <b>" + Формат(ВсегоОшибок, "ЧГ=0") + "</b>";
	Если ЗначениеЗаполнено(ИдентификаторПакета) Тогда
		HTML = HTML + " · Пакет: <code>" + Экр(ИдентификаторПакета) + "</code>";
	Иначе
		HTML = HTML + " · за " + Формат(ТекущаяДата(), "ДФ=dd.MM.yyyy") + " (последние 200)";
	КонецЕсли;
	HTML = HTML + "</div>";

	Если ВсегоОшибок = 0 Тогда
		HTML = HTML + "<div class=""err-empty"">✓ Ошибок нет</div></div>";
		Возврат ОбёрткаHTML(HTML);
	КонецЕсли;

	// 3. Группы по категориям с шаблонами рекомендаций
	Для Каждого КЗ Из ПоКатегориям Цикл
		Код = КЗ.Ключ;
		Ошибки = КЗ.Значение;
		Шаблон = ПолучитьШаблонКатегории(Код);

		HTML = HTML + "<div class=""err-cat err-cat-" + Шаблон.Класс + """>";
		HTML = HTML + "<div class=""err-cat-head"">"
			+ "<span class=""err-cat-icon"">" + Шаблон.Иконка + "</span>"
			+ "<span class=""err-cat-title"">" + Экр(Шаблон.Заголовок) + "</span>"
			+ "<span class=""err-cat-count"">" + Формат(Ошибки.Количество(), "ЧГ=0") + "</span>"
			+ "</div>";
		HTML = HTML + "<div class=""err-cat-action""><b>Что делать:</b> "
			+ Экр(Шаблон.ЧтоДелать);
		// H: ссылка на статью помощи по коду ошибки (если есть)
		КлючСтатьи = КлючСтатьиПомощиПоКатегории(Код);
		Если ЗначениеЗаполнено(КлючСтатьи) Тогда
			HTML = HTML + " <a href=""#"" data-action=""help|" + КлючСтатьи
				+ """ style=""color:#0d4a8a;font-weight:bold""> Подробнее</a>";
		КонецЕсли;
		HTML = HTML + "</div>";
		HTML = HTML + "<table class=""err-table""><tr>"
			+ "<th>Время</th><th>Пакет</th><th>Значение</th><th>Сообщение</th></tr>";
		Для Каждого О Из Ошибки Цикл
			HTML = HTML + "<tr>"
				+ "<td class=""err-time"">" + Формат(О.Время, "ДФ=dd.MM HH:mm") + "</td>"
				+ "<td class=""err-pack"">" + Экр(Лев(СокрЛП(О.Пакет), 40)) + "</td>"
				+ "<td class=""err-val""><code>" + Экр(Лев(СокрЛП(О.Значение), 40)) + "</code></td>"
				+ "<td>" + Экр(О.Сообщение) + "</td>"
				+ "</tr>";
		КонецЦикла;
		HTML = HTML + "</table></div>";
	КонецЦикла;

	HTML = HTML + "</div>";
	Возврат ОбёрткаHTML(HTML);

КонецФункции

// Стиль для HTML-отчёта ошибок (F).
//
Функция СтильОшибок()
	CSS = " .err-page{padding:8px;font-family:'Segoe UI',Tahoma,Arial,sans-serif;font-size:12px}"
		+ " .err-page h1{font-size:16px;margin:0 0 6px 0;color:#1a1a1a}"
		+ " .err-summary{color:#555;margin-bottom:12px;font-size:12px}"
		+ " .err-summary code{background:#f0f0f0;padding:1px 4px;border-radius:2px;font-size:11px}"
		+ " .err-empty{padding:20px;text-align:center;color:#1a7a30;background:#d4edda;border-radius:4px}"
		+ " .err-cat{margin-bottom:10px;border-radius:4px;border:1px solid #ccc;overflow:hidden}"
		+ " .err-cat-head{padding:6px 10px;display:flex;align-items:center;gap:8px;font-weight:bold}"
		+ " .err-cat-icon{font-size:16px}"
		+ " .err-cat-title{flex:1;font-size:13px}"
		+ " .err-cat-count{background:rgba(0,0,0,0.15);padding:1px 8px;border-radius:10px;font-size:11px}"
		+ " .err-cat-action{padding:6px 10px;background:#fff;border-top:1px solid #e0e0e0;font-size:12px;color:#333}"
		+ " .err-cat-server .err-cat-head{background:#fff3cd;color:#856404}"   // STS HTTP 5xx — жёлтый
		+ " .err-cat-auth .err-cat-head  {background:#f8d7da;color:#721c24}"   // авторизация — красный
		+ " .err-cat-net .err-cat-head   {background:#cfe2ff;color:#084298}"   // сеть — синий
		+ " .err-cat-mapping .err-cat-head{background:#e2d4f5;color:#3d2974}"  // маппинг — фиолетовый
		+ " .err-cat-nsi .err-cat-head   {background:#d1ecf1;color:#0c5460}"   // НСИ — бирюза
		+ " .err-cat-posting .err-cat-head{background:#f8d7da;color:#721c24}"  // проведение — красный
		+ " .err-cat-other .err-cat-head {background:#e2e3e5;color:#383d41}"   // прочее — серый
		+ " .err-table{width:100%;border-collapse:collapse;background:#fff}"
		+ " .err-table th{background:#f5f5f5;border:1px solid #e0e0e0;padding:4px 6px;font-size:11px;text-align:left}"
		+ " .err-table td{border:1px solid #f0f0f0;padding:4px 6px;font-size:11px;vertical-align:top}"
		+ " .err-time{white-space:nowrap;color:#666}"
		+ " .err-pack{font-family:'Consolas','Courier New',monospace;font-size:10px;color:#666}"
		+ " .err-val code{background:#f5f5f5;padding:1px 3px;font-size:10px}";
	Возврат "<style>" + CSS + "</style>";
КонецФункции

// Получить шаблон отображения для категории ошибки (F).
//
// Параметры:
//   Код - ПеречислениеСсылка.TL_КодОшибкиЗагрузки
//
// Возвращаемое значение:
//   Структура(Заголовок, Иконка, Класс, ЧтоДелать)
//
Функция ПолучитьШаблонКатегории(Код)

	Перечисления_ = Перечисления.TL_КодОшибкиЗагрузки;
	Заголовок = "Прочее"; Иконка = "❓"; Класс = "other";
	ЧтоДелать = "Откройте сообщение, посмотрите детали в журнале регистрации.";

	Если Код = Перечисления_.STS_ОшибкаСервера Тогда
		Заголовок = "Серверная ошибка STS (HTTP 5xx)";
		Иконка = "⚠"; Класс = "server";
		ЧтоДелать = "Это не настройка БП ГИГ. Подождите 15 минут и повторите загрузку."
			+ " Если ошибка повторится — обратитесь в техподдержку pos.autooplata.ru"
			+ " с номером смены и временем возникновения.";
	ИначеЕсли Код = Перечисления_.STS_ОшибкаАвторизации Тогда
		Заголовок = "Ошибка авторизации STS";
		Иконка = ""; Класс = "auth";
		ЧтоДелать = "Откройте: «TradeLedger» → «Настройки расширения»."
			+ " Проверьте логин и пароль в разделе «Подключение». Возможно истёк срок токена API.";
	ИначеЕсли Код = Перечисления_.STS_ОшибкаСети Тогда
		Заголовок = "Сетевая ошибка";
		Иконка = ""; Класс = "net";
		ЧтоДелать = "Проверьте подключение к интернету и VPN на сервере 1С."
			+ " Если связь восстановлена — повторите загрузку периода.";
	ИначеЕсли Код = Перечисления_.МаппингОплат_НеНайден Тогда
		Заголовок = "Не настроен маппинг видов оплат";
		Иконка = ""; Класс = "mapping";
		ЧтоДелать = "Откройте: «TradeLedger» → «Маппинг видов оплат STS»."
			+ " Создайте запись с Образцом (часть имени из STS) и Каналом"
			+ " (cards / online / ledger / voucher / retail_cash / retail_card / writeoff_fuel)."
			+ " До настройки незамапленные строки пропускаются без потерь.";
	ИначеЕсли Код = Перечисления_.НСИ_НеНайдена Тогда
		Заголовок = "НСИ не найдена";
		Иконка = ""; Класс = "nsi";
		ЧтоДелать = "Создайте недостающий элемент в БП ГИГ (Контрагенты / Номенклатура /"
			+ " Договоры / Склады) с ИНН/КПП или Кодом из колонки «Значение».";
	ИначеЕсли Код = Перечисления_.ВидОперацииНеСопоставлен Тогда
		Заголовок = "Вид операции не сопоставлен";
		Иконка = ""; Класс = "mapping";
		ЧтоДелать = "Источник прислал вид операции, который наша связка ещё не покрывает."
			+ " Передайте код вида операции (из колонки «Значение») разработчику.";
	ИначеЕсли Код = Перечисления_.СтавкаНДС_НеСопоставлена Тогда
		Заголовок = "Ставка НДС не определена";
		Иконка = ""; Класс = "nsi";
		ЧтоДелать = "Источник прислал неизвестную ставку НДС. Допустимые ставки прописываются"
			+ " в TL_МаппингЦБ.СтавкаНДСПоСтроке (без молчаливого 22%).";
	ИначеЕсли Код = Перечисления_.СчётУчёта_НеОпределён Тогда
		Заголовок = "Счёт учёта не определён";
		Иконка = ""; Класс = "nsi";
		ЧтоДелать = "В карточке номенклатуры или учётной политике не задан счёт учёта."
			+ " Откройте товар и заполните вкладку «Счета учёта»: 41.02 для розницы.";
	ИначеЕсли Код = Перечисления_.ОшибкаПроведения Тогда
		Заголовок = "Ошибка проведения документа";
		Иконка = ""; Класс = "posting";
		ЧтоДелать = "Проверьте остатки на складе, договоры контрагентов, заполненность"
			+ " обязательных реквизитов документа. Откройте документ из колонки и попробуйте провести вручную.";
	ИначеЕсли Код = Перечисления_.СуммаНеСходится Тогда
		Заголовок = "Сумма не сходится";
		Иконка = ""; Класс = "posting";
		ЧтоДелать = "Сумма в шапке не равна сумме строк. Проверьте источник пакета или"
			+ " пересчитайте суммы вручную.";
	ИначеЕсли Код = Перечисления_.ОбязательноеПолеПусто Тогда
		Заголовок = "Обязательное поле пустое";
		Иконка = ""; Класс = "nsi";
		ЧтоДелать = "Источник не передал обязательное поле. Проверьте источник пакета"
			+ " (имя поля — в колонке «Значение»).";
	ИначеЕсли Код = Перечисления_.Смена_БезПродаж Тогда
		Заголовок = "Смена без продаж — аномалия";
		Иконка = ""; Класс = "server";
		ЧтоДелать = "Реальных смен без продаж не бывает. В STS пришёл пустой sales[]"
			+ " или у всех строк volume=0. Возможные причины: смена ещё не закрыта,"
			+ " STS не получил данные от АЗС, тест-смена. Свяжитесь с оператором АЗС"
			+ " или техподдержкой pos.autooplata.ru.";
	ИначеЕсли Код = Перечисления_.Смена_БезРозницы Тогда
		Заголовок = "ℹ Смена без розничных продаж (только корпоратив)";
		Иконка = ""; Класс = "info";
		ЧтоДелать = "ОРП не создан — в смене продажи только через корпоратив"
			+ " (карты/талоны/ведомости). Перемещения на виртуальные склады созданы."
			+ " Если ожидалась розница — проверьте: (а) все ли service_code сопоставлены"
			+ " в TL_Настройки.НайтиНоменклатуруТоплива, (б) сопоставлен ли pay_type"
			+ " «Наличные»/«СберБанк» в TL_МаппингОплат.";
	ИначеЕсли Код = Перечисления_.Топливо_НеизвестныйКод Тогда
		Заголовок = "Неизвестный код топлива из STS";
		Иконка = "⛽"; Класс = "nsi";
		ЧтоДелать = "STS прислал service_code, которого нет в маппинге номенклатуры"
			+ " (обычно 1=АИ-92, 2=АИ-95, 3=ДТ). Возможно новое брендовое топливо"
			+ " (АИ-98, G-Drive). Добавьте сопоставление в TL_Настройки или в"
			+ " карточку Номенклатуры (реквизит «Код STS»). До этого литры в смене"
			+ " пропускаются без потери ОРП.";
	ИначеЕсли Код = Перечисления_.Документ_УжеЗагружен Тогда
		Заголовок = "ℹ Документ уже загружен (идемпотентность)";
		Иконка = "✓"; Класс = "info";
		ЧтоДелать = "Это не ошибка, а штатное поведение: документ с тем же"
			+ " КлючЗагрузки уже есть. Если нужна перезагрузка — сначала"
			+ " удалите документ через «Удалить выбранные» (он пометится"
			+ " на удаление), затем загружайте заново.";
	ИначеЕсли Код = Перечисления_.Переклассификация_Топлива Тогда
		Заголовок = "Переклассификация топлива — ручная обработка";
		Иконка = ""; Класс = "mapping";
		ЧтоДелать = "Найдена ТТН с отрицательными значениями. Это"
			+ " не реальное поступление от поставщика, а переклассификация"
			+ " номенклатуры на АЗС (например, ДТ зимний → ДТ обычный при"
			+ " сезонном переходе). В паре идёт положительная ТТН того же"
			+ " номера в соседней смене. Документ автоматически не создан."
			+ "<br><b>Действия бухгалтера:</b>"
			+ "<br>1. Документ.СписаниеТоваров на номенклатуру источника"
			+ " (например ДТ зим.) — на объём из колонки «Значение»."
			+ "<br>2. Документ.ОприходованиеТоваров на номенклатуру приёмника"
			+ " (например ДТ) — на тот же объём."
			+ "<br>Итог по складу = 0, переклассификация выполнена.";
	КонецЕсли;

	Возврат Новый Структура("Заголовок, Иконка, Класс, ЧтоДелать",
		Заголовок, Иконка, Класс, ЧтоДелать);

КонецФункции

// Полноэкранный отчёт = отчёт ошибок (даже если их нет) + крупный лог сессии.
// Открывается через ФормаДетали с ТипДокумента="ОтчётОшибок".
//
Функция СформироватьПолноэкранныйОтчётОшибокИЛог(ТекстЛога = "", МоментОт = Неопределено) Экспорт

	// Собираем HTML отчёта ошибок (внутренние стили + body)
	Заголовок = ?(ЗначениеЗаполнено(МоментОт),
		"Журнал ошибок текущей сессии (с " + Формат(МоментОт, "ДФ='HH:mm:ss'") + ")",
		"Журнал ошибок загрузки за сегодня");
	HTMLОшибок = СформироватьДетальныйОтчётОшибок("", Заголовок, МоментОт);
	// Из обёртки <!DOCTYPE…><body>…</body></html> вытащим только тело
	ТелоОшибок = ИзвлечьТелоИзHTML(HTMLОшибок);

	HTML = СтильПолноэкранногоОтчёта();
	HTML = HTML + "<div class=""rep-page"">";
	HTML = HTML + "<h1> Подробный отчёт ошибок и лог загрузки</h1>";
	HTML = HTML + "<div class=""rep-block"">" + ТелоОшибок + "</div>";

	HTML = HTML + "<h2> Лог текущей сессии</h2>";
	Если ПустаяСтрока(ТекстЛога) Тогда
		HTML = HTML + "<div class=""rep-empty"">Лог пустой — действий в этой сессии ещё не было.</div>";
	Иначе
		HTML = HTML + "<pre class=""rep-log"">" + Экр(ТекстЛога) + "</pre>";
	КонецЕсли;

	HTML = HTML + "</div>";
	Возврат ОбёрткаHTML(HTML);

КонецФункции

// Вытащить только то что внутри <body>…</body> из HTML-документа.
Функция ИзвлечьТелоИзHTML(HTML)
	НачBody = СтрНайти(НРег(HTML), "<body");
	Если НачBody = 0 Тогда
		Возврат HTML;
	КонецЕсли;
	НачКонца = СтрНайти(НРег(HTML), "</body");
	Если НачКонца = 0 Тогда
		НачКонца = СтрДлина(HTML);
	КонецЕсли;
	// Найти конец открывающего тега <body…>
	НачКонтента = СтрНайти(Сред(HTML, НачBody), ">");
	Если НачКонтента = 0 Тогда
		Возврат HTML;
	КонецЕсли;
	Старт = НачBody + НачКонтента;
	Длина = НачКонца - Старт;
	Если Длина < 1 Тогда
		Возврат HTML;
	КонецЕсли;
	Возврат Сред(HTML, Старт, Длина);
КонецФункции

// CSS для полноэкранного отчёта.
Функция СтильПолноэкранногоОтчёта()
	CSS = " .rep-page{padding:16px 24px;font-family:'Segoe UI',Tahoma,Arial,sans-serif;font-size:13px;color:#1a1a1a}"
		+ " .rep-page h1{font-size:22px;margin:0 0 14px 0;color:#0d4a8a;border-bottom:2px solid #d0d9e8;padding-bottom:8px}"
		+ " .rep-page h2{font-size:18px;margin:24px 0 10px 0;color:#1a4d7a}"
		+ " .rep-block{margin-bottom:24px}"
		+ " .rep-empty{padding:20px;text-align:center;color:#888;background:#f5f5f5;border-radius:4px}"
		+ " .rep-log{background:#fafaf5;color:#1a1a1a;padding:14px 18px;border:1px solid #d8d8d0;border-radius:4px;font-family:'Consolas','Courier New',monospace;font-size:13px;line-height:1.7;white-space:pre-wrap;word-wrap:break-word;max-height:600px;overflow-y:auto}";
	Возврат "<style>" + CSS + "</style>";
КонецФункции

// Получить ключ статьи помощи по категории ошибки (H-связка с F).
// Возвращает имя файла из docs/help/ (без .md) — клик на « Подробнее»
// в отчёте ошибок ведёт на соответствующую статью.
//
Функция КлючСтатьиПомощиПоКатегории(Код)
	Перечисления_ = Перечисления.TL_КодОшибкиЗагрузки;
	Если Код = Перечисления_.МаппингОплат_НеНайден Тогда
		Возврат "faq_документ_не_создался";
	ИначеЕсли Код = Перечисления_.НСИ_НеНайдена Тогда
		Возврат "faq_неизвестный_контрагент";
	ИначеЕсли Код = Перечисления_.ОшибкаПроведения Тогда
		Возврат "диагностика_ошибок";
	ИначеЕсли Код = Перечисления_.STS_ОшибкаСервера Тогда
		Возврат "checklist_реакция_на_ошибку";
	ИначеЕсли Код = Перечисления_.STS_ОшибкаАвторизации Тогда
		Возврат "checklist_первая_настройка";
	ИначеЕсли Код = Перечисления_.STS_ОшибкаСети Тогда
		Возврат "checklist_реакция_на_ошибку";
	ИначеЕсли Код = Перечисления_.СтавкаНДС_НеСопоставлена Тогда
		Возврат "проводки_покупки";
	ИначеЕсли Код = Перечисления_.СуммаНеСходится Тогда
		Возврат "faq_расхождение_в_сверке";
	ИначеЕсли Код = Перечисления_.ВидОперацииНеСопоставлен Тогда
		Возврат "gloss_kind";
	ИначеЕсли Код = Перечисления_.Смена_БезПродаж Тогда
		Возврат "checklist_реакция_на_ошибку";
	ИначеЕсли Код = Перечисления_.Смена_БезРозницы Тогда
		Возврат "gloss_коды_ошибок";
	ИначеЕсли Код = Перечисления_.Топливо_НеизвестныйКод Тогда
		Возврат "faq_номенклатура_не_найдена";
	ИначеЕсли Код = Перечисления_.Документ_УжеЗагружен Тогда
		Возврат "faq_уже_загружено";
	ИначеЕсли Код = Перечисления_.Переклассификация_Топлива Тогда
		Возврат "faq_переклассификация_топлива";
	КонецЕсли;
	Возврат "gloss_коды_ошибок";
КонецФункции

// ==========================================================================
//  ПОМОЩЬ — оглавление и статьи (H-направление плана luminous-humming-curry)
// ==========================================================================

// Сформировать HTML-оглавление встроенной помощи с группировкой по разделам.
// Контент — из общего модуля TL_ПомощьКонтент (автогенерируемый из docs/help).
//
Функция СформироватьОглавлениеПомощи() Экспорт

	Попытка
		Статьи = TL_ПомощьКонтент.ПолучитьВсеСтатьи();
	Исключение
		Возврат ОбёрткаHTML(СтильПомощи()
			+ "<div class=""help-page""><h1>Помощь TradeLedger</h1>"
			+ "<div class=""help-empty"">Контент помощи ещё не установлен. "
			+ "Пересоберите расширение или обратитесь к ELSY.</div></div>");
	КонецПопытки;

	// Группировка по полю Группа
	Группы = Новый Соответствие; // ИмяГруппы → Массив(Структура(Ключ, Заголовок))
	Для Каждого КЗ Из Статьи Цикл
		Ключ = КЗ.Ключ;
		Стр = КЗ.Значение;
		Гр = Стр.Группа;
		Список = Группы.Получить(Гр);
		Если Список = Неопределено Тогда
			Список = Новый Массив;
			Группы.Вставить(Гр, Список);
		КонецЕсли;
		Список.Добавить(Новый Структура("Ключ, Заголовок", Ключ, Стр.Заголовок));
	КонецЦикла;

	HTML = СтильПомощи() + "<div class=""help-page"">";
	HTML = HTML + "<h1> Помощь TradeLedger</h1>";
	HTML = HTML + "<div class=""help-intro"">Встроенная справка бухгалтеру. "
		+ "Кликни на статью — откроется содержимое. В каждой статье есть кнопка «← К оглавлению».</div>";

	ПорядокГрупп = Новый Массив;
	ПорядокГрупп.Добавить(Новый Структура("Имя, Заголовок, Иконка",
		"intro",      "С чего начать",            ""));
	ПорядокГрупп.Добавить(Новый Структура("Имя, Заголовок, Иконка",
		"checklist",  "Чек-листы",                "✅"));
	ПорядокГрупп.Добавить(Новый Структура("Имя, Заголовок, Иконка",
		"loading",    "Загрузка и переключение",  ""));
	ПорядокГрупп.Добавить(Новый Структура("Имя, Заголовок, Иконка",
		"accounting", "Проводки и учёт",          ""));
	ПорядокГрупп.Добавить(Новый Структура("Имя, Заголовок, Иконка",
		"food",       "Общепит",                  ""));
	ПорядокГрупп.Добавить(Новый Структура("Имя, Заголовок, Иконка",
		"diag",       "Сверка и ошибки",          ""));
	ПорядокГрупп.Добавить(Новый Структура("Имя, Заголовок, Иконка",
		"setup",      "Настройки",                "⚙"));
	ПорядокГрупп.Добавить(Новый Структура("Имя, Заголовок, Иконка",
		"faq",        "FAQ — частые вопросы",     "❓"));
	ПорядокГрупп.Добавить(Новый Структура("Имя, Заголовок, Иконка",
		"gloss",      "Словари и термины",        ""));
	ПорядокГрупп.Добавить(Новый Структура("Имя, Заголовок, Иконка",
		"process",    "Прочее",                   ""));

	Для Каждого ГрупИнфо Из ПорядокГрупп Цикл
		Список = Группы.Получить(ГрупИнфо.Имя);
		Если Список = Неопределено ИЛИ Список.Количество() = 0 Тогда
			Продолжить;
		КонецЕсли;
		HTML = HTML + "<div class=""help-section"">"
			+ "<h2><span class=""help-icon"">" + ГрупИнфо.Иконка + "</span> "
			+ Экр(ГрупИнфо.Заголовок) + " <span class=""help-count"">("
			+ Список.Количество() + ")</span></h2>";
		HTML = HTML + "<ul class=""help-list"">";
		Для Каждого С Из Список Цикл
			HTML = HTML + "<li><a href=""#"" data-action=""help|"
				+ С.Ключ + """>" + Экр(С.Заголовок) + "</a></li>";
		КонецЦикла;
		HTML = HTML + "</ul></div>";
	КонецЦикла;

	HTML = HTML + "</div>";
	Возврат ОбёрткаHTML(HTML);

КонецФункции

// Сформировать HTML конкретной статьи помощи.
//
// Параметры:
//   Ключ - Строка - идентификатор статьи (имя файла без .md)
//
Функция СформироватьСтатьюПомощи(Ключ) Экспорт

	Попытка
		Стр = TL_ПомощьКонтент.ПолучитьСтатью(Ключ);
	Исключение
		Стр = Неопределено;
	КонецПопытки;

	Если Стр = Неопределено Тогда
		HTML = СтильПомощи() + "<div class=""help-page"">"
			+ "<div class=""help-nav""><a href=""#"" data-action=""help|оглавление"">← К оглавлению</a></div>"
			+ "<h1>Статья не найдена</h1>"
			+ "<div class=""help-empty"">Статья с ключом «" + Экр(Ключ) + "» отсутствует.</div>"
			+ "</div>";
		Возврат ОбёрткаHTML(HTML);
	КонецЕсли;

	HTML = СтильПомощи() + "<div class=""help-page"">";
	HTML = HTML + "<div class=""help-nav""><a href=""#"" data-action=""help|оглавление"">← К оглавлению</a></div>";
	HTML = HTML + "<div class=""help-content"">" + МаркдаунВHTML(Стр.Markdown) + "</div>";
	HTML = HTML + "<div class=""help-nav""><a href=""#"" data-action=""help|оглавление"">← К оглавлению</a></div>";
	HTML = HTML + "</div>";
	Возврат ОбёрткаHTML(HTML);

КонецФункции

// CSS-стиль для страниц помощи (H).
//
Функция СтильПомощи()
	CSS = " .help-page{padding:12px 20px;font-family:'Segoe UI',Tahoma,Arial,sans-serif;font-size:13px;line-height:1.6;color:#1a1a1a;max-width:900px}"
		+ " .help-page h1{font-size:22px;margin:0 0 12px 0;color:#0d4a8a;border-bottom:2px solid #d0d9e8;padding-bottom:6px}"
		+ " .help-page h2{font-size:16px;margin:18px 0 8px 0;color:#1a4d7a}"
		+ " .help-page h3{font-size:14px;margin:14px 0 6px 0;color:#2a5d8a}"
		+ " .help-intro{color:#555;margin-bottom:14px;padding:8px 12px;background:#f5f9fc;border-left:3px solid #0d4a8a;border-radius:2px}"
		+ " .help-empty{padding:20px;text-align:center;color:#888;background:#f5f5f5;border-radius:4px}"
		+ " .help-section{margin-bottom:14px}"
		+ " .help-icon{font-size:18px;margin-right:4px}"
		+ " .help-count{color:#888;font-weight:normal;font-size:13px}"
		+ " .help-list{list-style:none;margin:0;padding:0 0 0 28px}"
		+ " .help-list li{padding:3px 0}"
		+ " .help-list a{color:#0d4a8a;text-decoration:none;cursor:pointer}"
		+ " .help-list a:hover{text-decoration:underline}"
		+ " .help-nav{margin:8px 0;font-size:12px}"
		+ " .help-nav a{color:#0d4a8a;text-decoration:none;cursor:pointer}"
		+ " .help-nav a:hover{text-decoration:underline}"
		+ " .help-content p{margin:8px 0}"
		+ " .help-content ul,.help-content ol{margin:6px 0 6px 24px;padding:0}"
		+ " .help-content li{margin:3px 0}"
		+ " .help-content code{background:#f0f0f0;padding:1px 5px;border-radius:2px;font-family:'Consolas','Courier New',monospace;font-size:12px;color:#a83232}"
		+ " .help-content pre{background:#f5f5f5;padding:8px 12px;border-radius:4px;overflow-x:auto;font-family:'Consolas','Courier New',monospace;font-size:12px}"
		+ " .help-content pre code{background:none;padding:0;color:#1a1a1a}"
		+ " .help-content strong{font-weight:bold;color:#000}"
		+ " .help-content em{font-style:italic}"
		+ " .help-content blockquote{margin:8px 0;padding:6px 12px;border-left:3px solid #c0c8d0;color:#555;background:#fafafa}"
		+ " .help-content table{border-collapse:collapse;margin:8px 0;font-size:12px}"
		+ " .help-content table th{background:#e8eef5;border:1px solid #c0c8d0;padding:4px 8px;text-align:left}"
		+ " .help-content table td{border:1px solid #d0d8e0;padding:4px 8px;vertical-align:top}"
		+ " .help-content hr{border:none;border-top:1px solid #d0d0d0;margin:12px 0}";
	Возврат "<style>" + CSS + "</style>";
КонецФункции

// Минимальный Markdown → HTML конвертер для статей помощи.
// Поддерживает: # заголовки, **жирный**, *курсив*, `код`, ```блок кода```,
//               - списки, 1. нумерация, > цитаты, --- разделитель,
//               пустая строка = новый абзац, простые таблицы (| col | col |).
//
Функция МаркдаунВHTML(Знач Текст) Экспорт

	Если ПустаяСтрока(Текст) Тогда
		Возврат "";
	КонецЕсли;

	Строки = СтрРазделить(Текст, Символы.ПС, Истина);
	Результат = "";
	ВБлокеКода   = Ложь;
	БлокКода     = "";
	ВСписке      = Ложь;
	ТипСписка    = "";  // "ul" или "ol"
	ВТаблице     = Ложь;
	ЗаголовокТаблицы = "";
	СтрокиТаблицы = Новый Массив;
	АбзацБуфер   = "";

	Для Каждого Стр Из Строки Цикл

		ОчищСтр = СокрЛП(Стр);

		// Блок кода ```
		Если СтрНачинаетсяС(ОчищСтр, "```") Тогда
			Если ВБлокеКода Тогда
				Результат = Результат + "<pre><code>" + Экр(БлокКода) + "</code></pre>";
				БлокКода = "";
				ВБлокеКода = Ложь;
			Иначе
				Результат = Результат + ЗакрытьОткрытые(АбзацБуфер, ВСписке, ТипСписка, ВТаблице, ЗаголовокТаблицы, СтрокиТаблицы);
				АбзацБуфер = ""; ВСписке = Ложь; ВТаблице = Ложь; СтрокиТаблицы = Новый Массив;
				ВБлокеКода = Истина;
			КонецЕсли;
			Продолжить;
		КонецЕсли;
		Если ВБлокеКода Тогда
			БлокКода = БлокКода + ?(ПустаяСтрока(БлокКода), "", Символы.ПС) + Стр;
			Продолжить;
		КонецЕсли;

		// Таблица: строка содержит | и не пустая
		Если СтрНайти(ОчищСтр, "|") > 0 И НЕ СтрНачинаетсяС(ОчищСтр, "    ") Тогда
			Если СтрНайти(ОчищСтр, "---") > 0 И ВТаблице Тогда
				// строка-разделитель шапки — пропускаем
				Продолжить;
			КонецЕсли;
			Если НЕ ВТаблице Тогда
				Результат = Результат + ЗакрытьОткрытые(АбзацБуфер, ВСписке, ТипСписка, Ложь, "", Новый Массив);
				АбзацБуфер = ""; ВСписке = Ложь;
				ВТаблице = Истина;
				ЗаголовокТаблицы = ОчищСтр;
			Иначе
				СтрокиТаблицы.Добавить(ОчищСтр);
			КонецЕсли;
			Продолжить;
		Иначе
			Если ВТаблице Тогда
				Результат = Результат + ВывестиТаблицу(ЗаголовокТаблицы, СтрокиТаблицы);
				ВТаблице = Ложь;
				ЗаголовокТаблицы = "";
				СтрокиТаблицы = Новый Массив;
			КонецЕсли;
		КонецЕсли;

		// Заголовки # ## ###
		Если СтрНачинаетсяС(ОчищСтр, "### ") Тогда
			Результат = Результат + ЗакрытьОткрытые(АбзацБуфер, ВСписке, ТипСписка, Ложь, "", Новый Массив);
			АбзацБуфер = ""; ВСписке = Ложь;
			Результат = Результат + "<h3>" + ОбработатьИнлайн(Сред(ОчищСтр, 5)) + "</h3>";
			Продолжить;
		КонецЕсли;
		Если СтрНачинаетсяС(ОчищСтр, "## ") Тогда
			Результат = Результат + ЗакрытьОткрытые(АбзацБуфер, ВСписке, ТипСписка, Ложь, "", Новый Массив);
			АбзацБуфер = ""; ВСписке = Ложь;
			Результат = Результат + "<h2>" + ОбработатьИнлайн(Сред(ОчищСтр, 4)) + "</h2>";
			Продолжить;
		КонецЕсли;
		Если СтрНачинаетсяС(ОчищСтр, "# ") Тогда
			Результат = Результат + ЗакрытьОткрытые(АбзацБуфер, ВСписке, ТипСписка, Ложь, "", Новый Массив);
			АбзацБуфер = ""; ВСписке = Ложь;
			Результат = Результат + "<h1>" + ОбработатьИнлайн(Сред(ОчищСтр, 3)) + "</h1>";
			Продолжить;
		КонецЕсли;

		// Разделитель ---
		Если ОчищСтр = "---" ИЛИ ОчищСтр = "***" Тогда
			Результат = Результат + ЗакрытьОткрытые(АбзацБуфер, ВСписке, ТипСписка, Ложь, "", Новый Массив);
			АбзацБуфер = ""; ВСписке = Ложь;
			Результат = Результат + "<hr/>";
			Продолжить;
		КонецЕсли;

		// Цитата >
		Если СтрНачинаетсяС(ОчищСтр, "> ") Тогда
			Результат = Результат + ЗакрытьОткрытые(АбзацБуфер, ВСписке, ТипСписка, Ложь, "", Новый Массив);
			АбзацБуфер = ""; ВСписке = Ложь;
			Результат = Результат + "<blockquote>" + ОбработатьИнлайн(Сред(ОчищСтр, 3)) + "</blockquote>";
			Продолжить;
		КонецЕсли;

		// Маркированный список -
		Если СтрНачинаетсяС(ОчищСтр, "- ") ИЛИ СтрНачинаетсяС(ОчищСтр, "* ") Тогда
			Если НЕ ВСписке ИЛИ ТипСписка <> "ul" Тогда
				Результат = Результат + ?(ВСписке, "</" + ТипСписка + ">", "") + ВыводитьАбзац(АбзацБуфер);
				АбзацБуфер = "";
				Результат = Результат + "<ul>";
				ВСписке = Истина; ТипСписка = "ul";
			КонецЕсли;
			Результат = Результат + "<li>" + ОбработатьИнлайн(Сред(ОчищСтр, 3)) + "</li>";
			Продолжить;
		КонецЕсли;

		// Нумерованный список 1.
		Если СтрДлина(ОчищСтр) > 2 И ИзЦифры(Лев(ОчищСтр, 1)) И НачинаетсяСНомера(ОчищСтр) Тогда
			Если НЕ ВСписке ИЛИ ТипСписка <> "ol" Тогда
				Результат = Результат + ?(ВСписке, "</" + ТипСписка + ">", "") + ВыводитьАбзац(АбзацБуфер);
				АбзацБуфер = "";
				Результат = Результат + "<ol>";
				ВСписке = Истина; ТипСписка = "ol";
			КонецЕсли;
			ПозТочки = СтрНайти(ОчищСтр, ". ");
			Результат = Результат + "<li>" + ОбработатьИнлайн(Сред(ОчищСтр, ПозТочки + 2)) + "</li>";
			Продолжить;
		КонецЕсли;

		// Пустая строка — конец абзаца / списка
		Если ПустаяСтрока(ОчищСтр) Тогда
			Если ВСписке Тогда
				Результат = Результат + "</" + ТипСписка + ">";
				ВСписке = Ложь;
			КонецЕсли;
			Если НЕ ПустаяСтрока(АбзацБуфер) Тогда
				Результат = Результат + "<p>" + ОбработатьИнлайн(АбзацБуфер) + "</p>";
				АбзацБуфер = "";
			КонецЕсли;
			Продолжить;
		КонецЕсли;

		// Обычный текст — копим в абзаце
		Если ВСписке Тогда
			Результат = Результат + "</" + ТипСписка + ">";
			ВСписке = Ложь;
		КонецЕсли;
		АбзацБуфер = АбзацБуфер + ?(ПустаяСтрока(АбзацБуфер), "", " ") + ОчищСтр;
	КонецЦикла;

	// Завершить открытые
	Если ВТаблице Тогда
		Результат = Результат + ВывестиТаблицу(ЗаголовокТаблицы, СтрокиТаблицы);
	КонецЕсли;
	Если ВСписке Тогда
		Результат = Результат + "</" + ТипСписка + ">";
	КонецЕсли;
	Если НЕ ПустаяСтрока(АбзацБуфер) Тогда
		Результат = Результат + "<p>" + ОбработатьИнлайн(АбзацБуфер) + "</p>";
	КонецЕсли;
	Если ВБлокеКода И НЕ ПустаяСтрока(БлокКода) Тогда
		Результат = Результат + "<pre><code>" + Экр(БлокКода) + "</code></pre>";
	КонецЕсли;

	Возврат Результат;

КонецФункции

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

// CSS + начало HTML для страницы деталей смены (стиль 1С из preview_html.html).
//
Функция СтильСмены()

	CSS = "* {margin:0;padding:0;box-sizing:border-box}"
		+ " body{background:#f5f0e0;color:#1a1a1a;font-family:'Segoe UI',Tahoma,Arial,sans-serif;font-size:12px;line-height:1.4;padding:8px}"
		+ " .header{background:#fff;border:1px solid #c0b898;padding:8px 12px;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center}"
		+ " .header-station{font-size:14px;font-weight:bold} .header-total{font-size:18px;font-weight:bold}"
		+ " .header-sub{color:#666;font-size:11px;text-align:right}"
		+ " table{width:100%;border-collapse:collapse;background:#fff;margin-bottom:6px;font-size:12px}"
		+ " table th{background:#e8deb8;border:1px solid #c0b898;padding:4px 8px;font-weight:bold;text-align:center;font-size:11px}"
		+ " table td{border:1px solid #d0c8a8;padding:3px 8px} table tr:hover{background:#faf6e8}"
		+ " .num{text-align:right;font-family:'Consolas','Courier New',monospace}"
		+ " .total-row{background:#f5f0d8;font-weight:bold} .total-row td{border-top:2px solid #c0b898}"
		+ " .negative{color:#c00}"
		+ " .section{background:#e8deb8;border:1px solid #c0b898;padding:4px 10px;font-weight:bold;font-size:12px;margin-bottom:1px}"
		+ " .provodka{font-size:10px;color:#666;font-family:'Consolas',monospace}"
		+ " .route-tag{font-size:10px;padding:1px 6px;border-radius:2px;font-weight:bold}"
		+ " .route-retail{background:#d4edda;color:#1a6b2a} .route-transfer{background:#fff3cd;color:#856404}"
		+ " .tanks-row{display:flex;gap:6px;margin-bottom:6px}"
		+ " .tank{flex:1;background:#fff;border:1px solid #c0b898;padding:6px 10px}"
		+ " .tank-head{display:flex;justify-content:space-between;margin-bottom:4px;font-weight:bold;font-size:11px}"
		+ " .tank-bar{height:14px;background:#e8e0c8;border:1px solid #d0c8a8;margin-bottom:4px;overflow:hidden;position:relative}"
		+ " .tank-bar-used{height:100%;position:absolute;left:0;top:0;opacity:0.25}"
		+ " .tank-bar-remain{height:100%;position:absolute;left:0;top:0}"
		+ " .fill-92{background:#c8a820} .fill-95{background:#3070c0} .fill-dt{background:#30a040}"
		+ " .tank-bar-label{position:absolute;top:0;height:100%;display:flex;align-items:center;font-size:9px;font-weight:bold;color:#555;padding:0 4px;font-family:'Consolas',monospace}"
		+ " .tank-bar-label.lbl-left{left:2px} .tank-bar-label.lbl-right{right:2px}"
		+ " .tank-stats td{border:none;padding:1px 4px;font-size:11px}"
		+ " .fuel-92{color:#8a6d00;font-weight:bold} .fuel-95{color:#1a5cb0;font-weight:bold} .fuel-dt{color:#1a7a30;font-weight:bold}"
		+ " .doc-row{display:flex;align-items:center;background:#fff;border:1px solid #d0c8a8;border-bottom:none;padding:5px 8px;gap:8px}"
		+ " .doc-row:last-child{border-bottom:1px solid #d0c8a8} .doc-row:hover{background:#faf6e8} .doc-row.disabled{color:#aaa}"
		+ " .doc-check{width:14px;height:14px;border:1px solid #999;background:#fff;display:flex;align-items:center;justify-content:center;font-size:11px;color:transparent;flex-shrink:0}"
		+ " .doc-row.checked .doc-check{background:#4a8c2a;border-color:#3a7020;color:#fff}"
		+ " .doc-name{flex:1;font-size:12px} .doc-detail{color:#555;font-size:10px;font-family:'Consolas',monospace}"
		+ " .doc-amount{font-family:'Consolas',monospace;font-weight:bold;font-size:12px;min-width:80px;text-align:right}";

	Возврат "<!DOCTYPE html><html lang=""ru""><head><meta charset=""UTF-8""><style>" + CSS + "</style></head><body>";

КонецФункции

// Обёртка HTML-документа (DOCTYPE + body со стилями).
//
Функция ОбёрткаHTML(Содержимое)
	Возврат "<!DOCTYPE html><html><head><meta charset=""utf-8""></head>"
		+ "<body style=""margin:0; padding:8px; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif;"
		+ " background:#F0F5FA;"">"
		+ Содержимое
		+ "</body></html>";
КонецФункции

// Мини-карточка дашборда (компактная, одна строка).
//
Функция МиниКарточка(Значение, Подпись, Цвет)
	Возврат "<div style=""width:120px; background:#fff; border-left:3px solid " + Цвет + ";"
		+ " padding:6px 10px; border-radius:8px; box-shadow:0 1px 2px rgba(0,0,0,0.06);"">"
		+ "<span style=""font-size:18px; font-weight:700; color:" + Цвет + ";"">" + Значение + "</span>"
		+ " <span style=""font-size:11px; color:#9CA3AF;"">" + Подпись + "</span>"
		+ "</div>";
КонецФункции

// Компактная карточка KPI для 4-колоночного дашборда.
//
Функция МиниКарточкаКомпакт(Значение, Подпись, Цвет)
	Возврат "<div style=""background:#F8FAFC; border-left:2px solid " + Цвет + ";"
		+ " padding:4px 8px; border-radius:4px; display:flex; align-items:center; gap:4px;"">"
		+ "<span style=""font-size:16px; font-weight:700; color:" + Цвет + ";"">" + Значение + "</span>"
		+ "<span style=""font-size:10px; color:#9CA3AF;"">" + Подпись + "</span>"
		+ "</div>";
КонецФункции

// Карточка дашборда (старая, для совместимости).
//
Функция Карточка(Значение, Подпись, Цвет, ФонЦвет)
	Возврат "<div style=""flex:1; background:#FFFFFF; border-left:4px solid " + Цвет + ";"
		+ " padding:12px; border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.08);"">"
		+ "<div style=""font-size:24px; font-weight:700; color:" + Цвет + ";"">" + Значение + "</div>"
		+ "<div style=""font-size:11px; color:#6B7280;"">" + Подпись + "</div>"
		+ "</div>";
КонецФункции

// Строка таблицы продаж (3 колонки).
//
Функция СтрокаТаблицы(Кол1, Кол2, Кол3, ЦветКол3)
	Возврат "<tr style=""border-bottom:1px solid #D9DFE7;"">"
		+ "<td style=""padding:4px 0;"">" + Кол1 + "</td>"
		+ "<td style=""text-align:right;"">" + Кол2 + "</td>"
		+ "<td style=""text-align:right; color:" + ЦветКол3 + ";"">" + Кол3 + "</td>"
		+ "</tr>";
КонецФункции

// Строка таблицы оплат (2 колонки).
//
Функция СтрокаТаблицы2(Кол1, Кол2)
	Возврат "<tr style=""border-bottom:1px solid #D9DFE7;"">"
		+ "<td style=""padding:4px 0;"">" + Кол1 + "</td>"
		+ "<td style=""text-align:right;"">" + Кол2 + "</td>"
		+ "</tr>";
КонецФункции

// Строка итогов (3 колонки).
//
Функция СтрокаИтого(Кол1, Кол2, Кол3)
	Возврат "<tr style=""font-weight:700; border-top:2px solid #1F2937;"">"
		+ "<td style=""padding:4px 0;"">" + Кол1 + "</td>"
		+ "<td style=""text-align:right;"">" + Кол2 + "</td>"
		+ "<td style=""text-align:right;"">" + Кол3 + "</td>"
		+ "</tr>";
КонецФункции

// Строка итогов (2 колонки).
//
Функция СтрокаИтого2(Кол1, Кол2)
	Возврат "<tr style=""font-weight:700; border-top:2px solid #1F2937;"">"
		+ "<td style=""padding:4px 0;"">" + Кол1 + "</td>"
		+ "<td style=""text-align:right;"">" + Кол2 + "</td>"
		+ "</tr>";
КонецФункции

// Карточка документа 1С (для панели деталей).
//
Функция КарточкаДокумента(Иконка, Текст, ФонЦвет, ЦветТекста)
	Возврат "<div style=""padding:6px 10px; margin:4px 0; background:" + ФонЦвет + ";"
		+ " border-radius:4px; font-size:13px;"">"
		+ "<span style=""color:" + ЦветТекста + ";"">" + Иконка + "</span> " + Текст
		+ "</div>";
КонецФункции

// Показатель (Масса/Плотность/Объём).
//
Функция Показатель(Подпись, Значение, Цвет)
	ЦветСтиль = ?(ЗначениеЗаполнено(Цвет), " color:" + Цвет + ";", "");
	Возврат "<div style=""text-align:center;"">"
		+ "<div style=""font-size:11px; color:#6B7280;"">" + Подпись + "</div>"
		+ "<div style=""font-size:20px; font-weight:700;" + ЦветСтиль + """>" + Значение + "</div>"
		+ "</div>";
КонецФункции

// Стрелка между показателями.
//
Функция Стрелка()
	Возврат "<div style=""font-size:24px; color:#D1D5DB;"">→</div>";
КонецФункции

// Шаг цепочки ТТН.
//
Функция ШагЦепочки(Номер, Заголовок, Описание, ФонЦвет, БордерЦвет)
	Возврат "<div style=""flex:1; padding:12px; background:" + ФонЦвет + "; border-radius:12px;"
		+ " border:1px solid " + БордерЦвет + "; text-align:center;"">"
		+ "<div style=""font-size:11px; color:#6B7280;"">" + Номер + "</div>"
		+ "<div style=""font-weight:600;"">" + Заголовок + "</div>"
		+ "<div style=""font-size:11px; color:#6B7280;"">" + Описание + "</div>"
		+ "</div>";
КонецФункции

// Собрать HTML-описание ТТН из данных БД (Перемещение + Комплектация).
// Используется при двойном клике на загруженную ТТН в форме TL_Загрузка —
// без обращения к STS API (быстро, ~100мс).
//
Функция СформироватьДеталиТТН_ИзБД(КлючБазовый, Заголовок) Экспорт

	СписокДок = TL_РегистрСтатусов.ПолучитьДокументыПоБазовомуКлючу(КлючБазовый);
	ПеремСсылка = Неопределено;
	КомплСсылка = Неопределено;
	Для Каждого ИнфоДок Из СписокДок Цикл
		Если Не ЗначениеЗаполнено(ИнфоДок.Ссылка) Тогда Продолжить; КонецЕсли;
		Если СтрНайти(ИнфоДок.Ключ, "|ПЕРЕМ") > 0 Тогда
			ПеремСсылка = ИнфоДок.Ссылка;
		ИначеЕсли СтрНайти(ИнфоДок.Ключ, "|КОМПЛ") > 0 Тогда
			КомплСсылка = ИнфоДок.Ссылка;
		КонецЕсли;
	КонецЦикла;
	Если Не ЗначениеЗаполнено(ПеремСсылка) Тогда
		Возврат ""; // в БД нет — fallback на STS
	КонецЕсли;

	// Перемещение: ТЧ Товары[0] — топливо в тоннах
	Запр = Новый Запрос;
	Запр.УстановитьПараметр("П", ПеремСсылка);
	Запр.Текст =
	"ВЫБРАТЬ ПЕРВЫЕ 1
	|	ТЧП.Номенклатура.Наименование КАК ИмяТопливаТонн,
	|	ТЧП.Количество                КАК МассаТонн
	|ИЗ Документ.ПеремещениеТоваров.Товары КАК ТЧП
	|ГДЕ ТЧП.Ссылка = &П";
	Вб = Запр.Выполнить().Выбрать();
	Если Вб.Следующий() Тогда
		ИмяТоплива = СокрЛП(Строка(Вб.ИмяТопливаТонн));
		МассаТонн  = Число(Вб.МассаТонн);
	Иначе
		Возврат "";
	КонецЕсли;

	// Комплектация: литры + плотность из Комментария
	Литры = 0;
	Плотность = 0;
	Если ЗначениеЗаполнено(КомплСсылка) Тогда
		Запр2 = Новый Запрос;
		Запр2.УстановитьПараметр("К", КомплСсылка);
		Запр2.Текст =
		"ВЫБРАТЬ ПЕРВЫЕ 1
		|	ТЧК.Количество       КАК Литры,
		|	Компл.Комментарий    КАК КомК
		|ИЗ Документ.КомплектацияНоменклатуры.Комплектующие КАК ТЧК
		|	ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.КомплектацияНоменклатуры КАК Компл
		|		ПО Компл.Ссылка = ТЧК.Ссылка
		|ГДЕ ТЧК.Ссылка = &К";
		Попытка
			Вб2 = Запр2.Выполнить().Выбрать();
			Если Вб2.Следующий() Тогда
				Литры = Число(Вб2.Литры);
				Ком = Строка(Вб2.КомК);
				Поз = СтрНайти(Ком, "плотность=");
				Если Поз > 0 Тогда
					Хвост = Сред(Ком, Поз + СтрДлина("плотность="));
					Попытка
						Плотность = Число(СтрЗаменить(СокрЛП(Хвост), ",", "."));
					Исключение
					КонецПопытки;
				КонецЕсли;
			КонецЕсли;
		Исключение
		КонецПопытки;
	КонецЕсли;
	МассаКг = МассаТонн * 1000;

	HTML = "<!DOCTYPE html><html><head><meta charset=""utf-8""><style>"
		+ "body{font-family:Segoe UI,sans-serif;font-size:13px;background:#F5F1E3;color:#1F2937;margin:0;padding:8px;}"
		+ ".header{background:#FFFCED;border:1px solid #D9D2B0;border-radius:6px;padding:14px 18px;margin-bottom:10px;"
		+ "display:flex;justify-content:space-between;align-items:flex-start;}"
		+ ".header h1{font-size:15px;margin:0;color:#1F2937;}"
		+ ".header .total{font-size:22px;font-weight:bold;color:#1F2937;}"
		+ ".header .sub{font-size:11px;color:#6B7280;margin-top:2px;}"
		+ ".section{background:#E6DDC1;font-weight:600;padding:5px 10px;border:1px solid #D9D2B0;margin-top:8px;}"
		+ "table{width:100%;border-collapse:collapse;background:#FFFCED;}"
		+ "td,th{padding:5px 10px;border:1px solid #D9D2B0;}"
		+ ".num{text-align:right;}"
		+ ".doc-row{display:flex;background:#FFFCED;border:1px solid #D9D2B0;padding:6px 10px;gap:10px;align-items:center;}"
		+ ".doc-row.checked{background:#E8F5E9;}"
		+ ".doc-check{color:#2E7D32;font-weight:bold;}"
		+ "</style></head><body>";

	HTML = HTML + "<div class=""header""><div>"
		+ "<h1>" + Экр(Заголовок) + "</h1></div><div style=""text-align:right;"">"
		+ "<div class=""total"">" + Формат(МассаТонн, "ЧДЦ=3") + " т</div>"
		+ "<div class=""sub"">" + Экр(ИмяТоплива) + " &middot; &rho;="
		+ ?(Плотность > 0, Формат(Плотность, "ЧДЦ=4"), "—") + "</div></div></div>";

	HTML = HTML + "<div class=""section"">Данные ТТН (из БД)</div>"
		+ "<table><tr><th>Параметр</th><th>Значение</th></tr>"
		+ "<tr><td>Топливо</td><td><b>" + Экр(ИмяТоплива) + "</b></td></tr>"
		+ "<tr><td>Масса</td><td class=""num""><b>" + Формат(МассаТонн, "ЧДЦ=3") + " т</b> ("
		+ Формат(МассаКг, "ЧДЦ=2") + " кг)</td></tr>"
		+ "<tr><td>Объём (книжный)</td><td class=""num""><b>" + Формат(Литры, "ЧДЦ=2") + " л</b></td></tr>"
		+ "<tr><td>Плотность</td><td class=""num"">"
		+ ?(Плотность > 0, Формат(Плотность, "ЧДЦ=4") + " кг/л", "&mdash;") + "</td></tr>"
		+ "</table>";

	HTML = HTML + "<div class=""section"">Документы (созданы)</div><div>";
	UUIDПерем = XMLСтрока(ПеремСсылка.УникальныйИдентификатор());
	HTML = HTML + "<div class=""doc-row checked""><div class=""doc-check"">&#10003;</div>"
		+ "<div><b>Перемещение</b> Осн.склад &rarr; АЗС &middot; "
		+ Формат(МассаТонн, "ЧДЦ=3") + " т"
		+ " <a href=""#"" data-action=""open|Перемещение|" + UUIDПерем
		+ """ style=""margin-left:8px;font-size:11px;""> открыть</a></div></div>";
	Если ЗначениеЗаполнено(КомплСсылка) Тогда
		UUIDКомпл = XMLСтрока(КомплСсылка.УникальныйИдентификатор());
		HTML = HTML + "<div class=""doc-row checked""><div class=""doc-check"">&#10003;</div>"
			+ "<div><b>Комплектация</b> тонны &rarr; литры &middot; "
			+ Формат(МассаТонн, "ЧДЦ=3") + " т &rarr; " + Формат(Литры, "ЧДЦ=2") + " л"
			+ " <a href=""#"" data-action=""open|Комплектация|" + UUIDКомпл
			+ """ style=""margin-left:8px;font-size:11px;""> открыть</a></div></div>";
	КонецЕсли;
	HTML = HTML + "</div></body></html>";
	Возврат HTML;

КонецФункции

// Сформировать HTML-карточку СМЕНЫ сопутки/общепита прямо из данных БП ГИГ.
//
// В отличие от СформироватьДеталиПакетаСмены (читает JSON-файл пакета, которого
// в витрине канала уже нет) — собирает картину из документов БД: по одному
// документу смены находит её координаты в РегистрСведений.TL_СоответствиеИсточников
// (КодАЗС + НомерСмены + ДатаСмены), собирает ВСЕ документы смены, их ТЧ Товары,
// оплаты ОРП и ФАКТИЧЕСКИЕ проводки из регистра бухгалтерии. Визуальный язык —
// тот же, что у топливной карточки смены (СтильСмены): шапка-итог, список
// документов со ссылками «открыть», сводная таблица товаров, проводки.
//
// Параметры:
//   Документ  - ДокументСсылка - любой документ смены (из строки витрины канала)
//   Заголовок - Строка         - готовый заголовок ("Смена N, АЗС K, дата")
//   Режим     - Строка         - "Сопутка" | "Общепит" (для подписи шапки)
//
// Возвращаемое значение:
//   Строка - готовый HTML-документ для ПолеHTML.
//
Функция СформироватьДеталиСменыСопуткиИзБД(Документ, Заголовок, Режим = "Сопутка") Экспорт

	Если Не ЗначениеЗаполнено(Документ) Тогда
		Возврат ОбёрткаHTML("<p style=""color:#888;"">Документ не указан</p>");
	КонецЕсли;

	// ===== 1. КООРДИНАТЫ СМЕНЫ ПО ДОКУМЕНТУ =====
	ЗапросС = Новый Запрос;
	ЗапросС.УстановитьПараметр("Док", Документ);
	ЗапросС.Текст =
	"ВЫБРАТЬ ПЕРВЫЕ 1
	|	СИ.КодАЗС      КАК КодАЗС,
	|	СИ.НомерСмены  КАК НомерСмены,
	|	СИ.ДатаСмены   КАК ДатаСмены
	|ИЗ РегистрСведений.TL_СоответствиеИсточников КАК СИ
	|ГДЕ СИ.СсылкаОбъект = &Док";
	ВыбС = ЗапросС.Выполнить().Выбрать();
	КодАЗС = 0; НомерСмены = ""; ДатаСмены = '00010101';
	Если ВыбС.Следующий() Тогда
		КодАЗС     = ?(ЗначениеЗаполнено(ВыбС.КодАЗС), ВыбС.КодАЗС, 0);
		НомерСмены = СокрЛП(Строка(ВыбС.НомерСмены));
		ДатаСмены  = ВыбС.ДатаСмены;
	КонецЕсли;

	// ===== 2. ВСЕ ДОКУМЕНТЫ СМЕНЫ (или один — если смена не определена) =====
	ДокументыСмены = Новый Массив;
	Если ЗначениеЗаполнено(НомерСмены) Тогда
		ЗапросД = Новый Запрос;
		ЗапросД.УстановитьПараметр("КодАЗС",     КодАЗС);
		ЗапросД.УстановитьПараметр("НомерСмены", НомерСмены);
		ЗапросД.УстановитьПараметр("ДатаСмены",  ДатаСмены);
		ЗапросД.Текст =
		"ВЫБРАТЬ РАЗЛИЧНЫЕ
		|	СИ.СсылкаОбъект КАК Документ
		|ИЗ РегистрСведений.TL_СоответствиеИсточников КАК СИ
		|ГДЕ СИ.КодАЗС = &КодАЗС
		|	И СИ.НомерСмены = &НомерСмены
		|	И СИ.ДатаСмены = &ДатаСмены";
		ВыбД = ЗапросД.Выполнить().Выбрать();
		Пока ВыбД.Следующий() Цикл
			Если ЗначениеЗаполнено(ВыбД.Документ) Тогда
				ДокументыСмены.Добавить(ВыбД.Документ);
			КонецЕсли;
		КонецЦикла;
	КонецЕсли;
	Если ДокументыСмены.Количество() = 0 Тогда
		ДокументыСмены.Добавить(Документ);
	КонецЕсли;

	// ===== 3. ОСНОВНОЙ ДОКУМЕНТ (тот, на который кликнули в витрине) =====
	ИмяМД = ""; Попытка ИмяМД = Документ.Метаданные().Имя; Исключение КонецПопытки;
	ТипЧит  = _ДетБД_ИмяТипаДок(ИмяМД);
	ЭтоИнвентаризация = (ИмяМД = "ИнвентаризацияТоваровНаСкладе");
	ЭтоОРП            = (ИмяМД = "ОтчетОРозничныхПродажах");

	НомерОсн = ""; Попытка НомерОсн = СокрЛП(Документ.Номер); Исключение КонецПопытки;
	ДатаОсн = '00010101'; Попытка ДатаОсн = Документ.Дата; Исключение КонецПопытки;
	СуммаОсн = 0; Попытка СуммаОсн = Документ.СуммаДокумента; Исключение КонецПопытки;
	ПроведёнОсн = Ложь; Попытка ПроведёнОсн = Документ.Проведен; Исключение КонецПопытки;

	// Заголовок карточки = ИМЕННО этот документ + ссылка на смену
	ЗаголовокДок = ТипЧит
		+ ?(ЗначениеЗаполнено(НомерОсн), " №" + НомерОсн, "")
		+ " · смена " + ?(ЗначениеЗаполнено(НомерСмены), НомерСмены, "?")
		+ ", АЗС " + XMLСтрока(КодАЗС)
		+ ?(ЗначениеЗаполнено(ДатаОсн) И ДатаОсн <> '00010101', " · " + Формат(ДатаОсн, "ДФ='dd.MM.yyyy'"), "");

	// ТЧ Товары + Оплаты ТОЛЬКО основного документа
	СтрокиТоваров = Новый Массив;
	СтрокиОплат   = Новый Массив;

	// Раскладка блюд (ингредиент→блюдо) из регистра — для группировки общепита под блюдом
	РаскладкаКэш = Новый Соответствие;
	Попытка
		ЗапросРБ = Новый Запрос;
		ЗапросРБ.УстановитьПараметр("Док", Документ);
		ЗапросРБ.Текст =
		"ВЫБРАТЬ РБ.Ингредиент КАК Ингредиент, РБ.Блюдо КАК Блюдо
		|ИЗ РегистрСведений.TL_РаскладкаБлюд КАК РБ
		|ГДЕ РБ.Документ = &Док";
		ВыбРБ = ЗапросРБ.Выполнить().Выбрать();
		Пока ВыбРБ.Следующий() Цикл
			РаскладкаКэш.Вставить(ВыбРБ.Ингредиент, ВыбРБ.Блюдо);
		КонецЦикла;
	Исключение
	КонецПопытки;

	ДокОб = Неопределено;
	Попытка ДокОб = Документ.ПолучитьОбъект(); Исключение КонецПопытки;
	Если ДокОб <> Неопределено Тогда
		ТЧТовары = Неопределено;
		Попытка ТЧТовары = ДокОб.Товары; Исключение ТЧТовары = Неопределено; КонецПопытки;
		Если ТЧТовары <> Неопределено Тогда
			Для Каждого Ст Из ТЧТовары Цикл
				СтТов = Новый Структура("Номенклатура, Количество, КоличествоУчет, Цена, Сумма, СтавкаНДС, СуммаНДС, Блюдо");
				СтТов.Номенклатура = ""; Попытка СтТов.Номенклатура = СокрЛП(Строка(Ст.Номенклатура)); Исключение КонецПопытки;
				СтТов.Количество = 0; Попытка СтТов.Количество = Ст.Количество; Исключение КонецПопытки;
				СтТов.КоличествоУчет = 0; Попытка СтТов.КоличествоУчет = Ст.КоличествоУчет; Исключение КонецПопытки;
				СтТов.Цена = 0; Попытка СтТов.Цена = Ст.Цена; Исключение КонецПопытки;
				СтТов.Сумма = 0; Попытка СтТов.Сумма = Ст.Сумма; Исключение КонецПопытки;
				СтТов.СтавкаНДС = ""; Попытка СтТов.СтавкаНДС = СокрЛП(Строка(Ст.СтавкаНДС)); Исключение КонецПопытки;
				СтТов.СуммаНДС = 0; Попытка СтТов.СуммаНДС = Ст.СуммаНДС; Исключение КонецПопытки;
				// Блюдо общепита: связь ингредиент→блюдо из регистра TL_РаскладкаБлюд
				СтТов.Блюдо = Неопределено;
				Попытка СтТов.Блюдо = РаскладкаКэш.Получить(Ст.Номенклатура); Исключение КонецПопытки;
				СтрокиТоваров.Добавить(СтТов);
			КонецЦикла;
		КонецЕсли;
		Если ЭтоОРП Тогда
			ТЧОпл = Неопределено;
			Попытка ТЧОпл = ДокОб.Оплата; Исключение ТЧОпл = Неопределено; КонецПопытки;
			Если ТЧОпл <> Неопределено Тогда
				Для Каждого Ст Из ТЧОпл Цикл
					СтОпл = Новый Структура("ВидОплаты, Сумма");
					СтОпл.ВидОплаты = ""; Попытка СтОпл.ВидОплаты = СокрЛП(Строка(Ст.ВидОплаты)); Исключение КонецПопытки;
					СтОпл.Сумма = 0; Попытка СтОпл.Сумма = Ст.СуммаОплаты; Исключение КонецПопытки;
					СтрокиОплат.Добавить(СтОпл);
				КонецЦикла;
			КонецЕсли;
			// Наличные = СуммаДокумента − Σ безналичных. ТЧ «Оплата» ОРП хранит только
			// безнал (карты/сертификаты); наличная выручка идёт на 50.01 отдельно. Без
			// этой строки блок «Оплаты» не сходится с выручкой.
			СуммаБезнал = 0;
			Для Каждого СтОпл Из СтрокиОплат Цикл
				СуммаБезнал = СуммаБезнал + СтОпл.Сумма;
			КонецЦикла;
			НаличныеСумма = СуммаОсн - СуммаБезнал;
			Если НаличныеСумма > 0.005 Тогда
				СтН = Новый Структура("ВидОплаты, Сумма", "Наличные", НаличныеСумма);
				СтрокиОплат.Вставить(0, СтН);
			КонецЕсли;
		КонецЕсли;
	КонецЕсли;

	// ===== 4. ПРОВОДКИ ТОЛЬКО ОСНОВНОГО ДОКУМЕНТА =====
	СтрокиПроводок = Новый Массив;
	ИтогоПроводки = 0;
	ЗапросП = Новый Запрос;
	ЗапросП.УстановитьПараметр("Док", Документ);
	ЗапросП.Текст =
	"ВЫБРАТЬ
	|	Хоз.СчетДт.Код КАК КодДт,
	|	Хоз.СчетКт.Код КАК КодКт,
	|	СУММА(Хоз.Сумма) КАК Сумма
	|ИЗ РегистрБухгалтерии.Хозрасчетный КАК Хоз
	|ГДЕ Хоз.Регистратор = &Док
	|СГРУППИРОВАТЬ ПО Хоз.СчетДт.Код, Хоз.СчетКт.Код
	|УПОРЯДОЧИТЬ ПО Сумма УБЫВ";
	Попытка
		ВыбП = ЗапросП.Выполнить().Выбрать();
		Пока ВыбП.Следующий() Цикл
			СтП = Новый Структура("Дт, Кт, Сумма");
			СтП.Дт = СокрЛП(Строка(ВыбП.КодДт));
			СтП.Кт = СокрЛП(Строка(ВыбП.КодКт));
			СтП.Сумма = ВыбП.Сумма;
			СтрокиПроводок.Добавить(СтП);
			ИтогоПроводки = ИтогоПроводки + ВыбП.Сумма;
		КонецЦикла;
	Исключение
	КонецПопытки;

	// ===== 5. ГЕНЕРАЦИЯ HTML =====
	HTML = СтильСмены();

	// --- ШАПКА: основной документ ---
	HTML = HTML + "<div class=""header""><div>"
		+ "<div class=""header-station"">" + Экр(Режим) + " &middot; " + Экр(ЗаголовокДок) + "</div>"
		+ "</div><div>"
		+ "<div class=""header-total"">"
		+ ?(СуммаОсн <> 0, Формат(СуммаОсн, "ЧДЦ=2; ЧРД=,; ЧГ=' '") + " &#8381;", "&mdash;") + "</div>"
		+ "<div class=""header-sub"">" + Экр(ТипЧит)
		+ " &middot; " + ?(ПроведёнОсн, "проведён", "не проведён")
		+ " &middot; " + XMLСтрока(СтрокиТоваров.Количество()) + " позиций</div>"
		+ "</div></div>";

	// --- НАВИГАЦИЯ: другие документы этой смены (текущий выделен) ---
	Если ДокументыСмены.Количество() > 1 Тогда
		HTML = HTML + "<div class=""section"">Документы смены (" + XMLСтрока(ДокументыСмены.Количество()) + ")</div><div>";
		Для Каждого Док Из ДокументыСмены Цикл
			ИмяМД2 = ""; Попытка ИмяМД2 = Док.Метаданные().Имя; Исключение КонецПопытки;
			ТипЧит2  = _ДетБД_ИмяТипаДок(ИмяМД2);
			ВидОткр2 = _ДетБД_ВидДляОткрытия(ИмяМД2);
			Бейдж2   = _ДетБД_БейджТипа(ИмяМД2);
			Номер2 = ""; Попытка Номер2 = СокрЛП(Док.Номер); Исключение КонецПопытки;
			Сумма2 = 0; Попытка Сумма2 = Док.СуммаДокумента; Исключение КонецПопытки;
			Проведён2 = Ложь; Попытка Проведён2 = Док.Проведен; Исключение КонецПопытки;
			UUID2 = ""; Попытка UUID2 = XMLСтрока(Док.УникальныйИдентификатор()); Исключение КонецПопытки;
			ЭтоТекущий = (Док = Документ);

			Если ЭтоТекущий Тогда
				ДействиеСсылка = " <span style=""font-size:11px;color:#4a8c2a;font-weight:bold;"">&#9679; открыт</span>";
			ИначеЕсли ЗначениеЗаполнено(UUID2) И ЗначениеЗаполнено(ВидОткр2) Тогда
				ДействиеСсылка = " <a href=""#"" data-action=""open|" + ВидОткр2 + "|" + UUID2
					+ """ style=""margin-left:8px;font-size:11px;text-decoration:none;""> открыть</a>";
			Иначе
				ДействиеСсылка = "";
			КонецЕсли;

			HTML = HTML + "<div class=""" + ?(Проведён2, "doc-row checked", "doc-row") + """"
				+ ?(ЭтоТекущий, " style=""background:#fffae0;border-left:3px solid #4a8c2a""", "") + ">"
				+ "<div class=""doc-name"">" + Бейдж2 + " <b>" + Экр(ТипЧит2) + "</b>"
				+ ?(ЗначениеЗаполнено(Номер2), " <span style=""color:#1F2937;"">№" + Экр(Номер2) + "</span>", "")
				+ ДействиеСсылка + "</div>"
				+ "<div class=""doc-amount"">" + ?(Сумма2 <> 0, Формат(Сумма2, "ЧДЦ=2"), "&mdash;") + "</div>"
				+ "</div>";
		КонецЦикла;
		HTML = HTML + "</div>";
	КонецЕсли;

	// --- ТОВАРЫ / ВЕДОМОСТЬ — основного документа ---
	HTML = HTML + "<div class=""section"" style=""margin-top:6px"">"
		+ ?(ЭтоИнвентаризация, "Ведомость инвентаризации", "Товары")
		+ " (" + XMLСтрока(СтрокиТоваров.Количество()) + ")</div>";
	Если СтрокиТоваров.Количество() = 0 Тогда
		HTML = HTML + "<div class=""doc-row disabled""><div class=""doc-name"">Товарных строк нет</div></div>";
	ИначеЕсли ЭтоИнвентаризация Тогда
		// Инвентаризация: Факт / Учёт / Отклонение (пустой «Факт» = по позиции нет фактического остатка)
		HTML = HTML + "<table><tr><th style=""width:24px"">#</th>"
			+ "<th>Номенклатура</th><th>Факт</th><th>Учёт</th><th>Откл.</th></tr>";
		Ном = 0;
		Для Каждого СтТов Из СтрокиТоваров Цикл
			Ном = Ном + 1;
			Откл = СтТов.Количество - СтТов.КоличествоУчет;
			ЦветОткл = ?(Откл = 0, "#999", ?(Откл > 0, "#1a7a30", "#c00"));
			HTML = HTML + "<tr><td class=""num"">" + XMLСтрока(Ном) + "</td>"
				+ "<td>" + Экр(СтТов.Номенклатура) + "</td>"
				+ "<td class=""num"">" + ?(СтТов.Количество <> 0, Формат(СтТов.Количество, "ЧДЦ=3"), "&mdash;") + "</td>"
				+ "<td class=""num"" style=""color:#888"">" + ?(СтТов.КоличествоУчет <> 0, Формат(СтТов.КоличествоУчет, "ЧДЦ=3"), "&mdash;") + "</td>"
				+ "<td class=""num"" style=""color:" + ЦветОткл + ";font-weight:bold"">"
				+ ?(Откл = 0, "0", ?(Откл > 0, "+", "") + Формат(Откл, "ЧДЦ=3")) + "</td></tr>";
		КонецЦикла;
		HTML = HTML + "</table>";
	ИначеЕсли ЭтоОРП Тогда
		// ОРП: блюда общепита раскрыты на ингредиенты (вариант B). Строки одного блюда
		// несут Спецификацию (ТТК), её Владелец = блюдо — группируем под именем блюда
		// (« Кофе (раскладка: сироп, кофе, молоко…)»). Строки без блюда — обычная сопутка.
		// Fallback на старые данные (без Спецификации): головная с ценой + нулевые за ней.
		Группы = Новый Массив;
		ТекГруппа = Неопределено;
		Для Каждого СтТов Из СтрокиТоваров Цикл
			ЕстьБлюдо = ЗначениеЗаполнено(СтТов.Блюдо);
			ЭтоНулевая = (СтТов.Сумма = 0 И СтТов.Цена = 0);
			Если ЕстьБлюдо Тогда
				Если ТекГруппа = Неопределено ИЛИ Не ЗначениеЗаполнено(ТекГруппа.Блюдо)
						ИЛИ ТекГруппа.Блюдо <> СтТов.Блюдо Тогда
					ТекГруппа = Новый Структура("Блюдо, ИмяБлюда, Строки",
						СтТов.Блюдо, СокрЛП(Строка(СтТов.Блюдо)), Новый Массив);
					Группы.Добавить(ТекГруппа);
				КонецЕсли;
				ТекГруппа.Строки.Добавить(СтТов);
			ИначеЕсли Не ЭтоНулевая ИЛИ ТекГруппа = Неопределено Тогда
				ТекГруппа = Новый Структура("Блюдо, ИмяБлюда, Строки", Неопределено, "", Новый Массив);
				ТекГруппа.Строки.Добавить(СтТов);
				Группы.Добавить(ТекГруппа);
			Иначе
				ТекГруппа.Строки.Добавить(СтТов);
			КонецЕсли;
		КонецЦикла;

		HTML = HTML + "<table><tr><th style=""width:24px"">#</th>"
			+ "<th>Номенклатура</th><th>Кол</th><th>Цена</th><th>Сумма</th><th>НДС</th><th>&sum; НДС</th></tr>";
		Ном = 0; ИтогоСумма = 0; ИтогоНДС = 0;
		Для Каждого Гр Из Группы Цикл
			Ном = Ном + 1;
			ЭтоБлюдо = (Гр.Строки.Количество() > 1) ИЛИ ЗначениеЗаполнено(Гр.Блюдо);
			ВыручкаГр = 0; НДСГр = 0;
			Для Каждого С Из Гр.Строки Цикл ВыручкаГр = ВыручкаГр + С.Сумма; НДСГр = НДСГр + С.СуммаНДС; КонецЦикла;
			ИтогоСумма = ИтогоСумма + ВыручкаГр;
			ИтогоНДС = ИтогоНДС + НДСГр;

			Если ЭтоБлюдо Тогда
				// заголовок блюда: имя блюда (или, если данных нет, первого ингредиента) + выручка группы
				ИмяГоловы = ?(ЗначениеЗаполнено(Гр.Блюдо), Гр.ИмяБлюда, СокрЛП(Строка(Гр.Строки[0].Номенклатура)));
				HTML = HTML + "<tr style=""background:#fbf3e0"">"
					+ "<td class=""num"">" + XMLСтрока(Ном) + "</td>"
					+ "<td> <b>" + Экр(ИмяГоловы) + "</b>"
					+ " <span style=""color:#b8860b;font-size:10px"">(раскладка &middot; "
					+ XMLСтрока(Гр.Строки.Количество()) + " ингр.)</span></td>"
					+ "<td class=""num"">&mdash;</td>"
					+ "<td class=""num"">&mdash;</td>"
					+ "<td class=""num""><b>" + Формат(ВыручкаГр, "ЧДЦ=2") + "</b></td>"
					+ "<td></td>"
					+ "<td class=""num"">" + ?(НДСГр <> 0, Формат(НДСГр, "ЧДЦ=2"), "&mdash;") + "</td></tr>";
				Для Каждого Инг Из Гр.Строки Цикл
					HTML = HTML + "<tr style=""color:#8a8a6a;font-size:11px"">"
						+ "<td></td>"
						+ "<td style=""padding-left:20px"">&#8627; " + Экр(Инг.Номенклатура) + "</td>"
						+ "<td class=""num"">" + Формат(Инг.Количество, "ЧДЦ=3") + "</td>"
						+ "<td class=""num"">&mdash;</td><td class=""num"">&mdash;</td>"
						+ "<td></td><td></td></tr>";
				КонецЦикла;
			Иначе
				// обычный товар сопутки (одна строка)
				Т = Гр.Строки[0];
				HTML = HTML + "<tr><td class=""num"">" + XMLСтрока(Ном) + "</td>"
					+ "<td>" + Экр(Т.Номенклатура) + "</td>"
					+ "<td class=""num"">" + Формат(Т.Количество, "ЧДЦ=3") + "</td>"
					+ "<td class=""num"">" + ?(Т.Цена <> 0, Формат(Т.Цена, "ЧДЦ=2"), "&mdash;") + "</td>"
					+ "<td class=""num"">" + ?(Т.Сумма <> 0, Формат(Т.Сумма, "ЧДЦ=2"), "&mdash;") + "</td>"
					+ "<td>" + Экр(Т.СтавкаНДС) + "</td>"
					+ "<td class=""num"">" + ?(Т.СуммаНДС <> 0, Формат(Т.СуммаНДС, "ЧДЦ=2"), "&mdash;") + "</td></tr>";
			КонецЕсли;
		КонецЦикла;
		HTML = HTML + "<tr class=""total-row""><td colspan=""4"">ИТОГО выручка</td>"
			+ "<td class=""num"">" + Формат(ИтогоСумма, "ЧДЦ=2") + "</td><td></td>"
			+ "<td class=""num"">" + Формат(ИтогоНДС, "ЧДЦ=2") + "</td></tr></table>";
		HTML = HTML + "<div style=""font-size:10px;color:#999;margin-top:2px"">"
			+ " — блюдо общепита (раскладка по ТТК), &#8627; — ингредиент (списание сырья)</div>";
	Иначе
		// Поступление / прочее: Кол / Цена / Сумма / НДС
		HTML = HTML + "<table><tr><th style=""width:24px"">#</th>"
			+ "<th>Номенклатура</th><th>Кол</th><th>Цена</th><th>Сумма</th><th>НДС</th><th>&sum; НДС</th></tr>";
		Ном = 0; ИтогоСумма = 0; ИтогоНДС = 0;
		Для Каждого СтТов Из СтрокиТоваров Цикл
			Ном = Ном + 1;
			ИтогоСумма = ИтогоСумма + СтТов.Сумма;
			ИтогоНДС = ИтогоНДС + СтТов.СуммаНДС;
			HTML = HTML + "<tr><td class=""num"">" + XMLСтрока(Ном) + "</td>"
				+ "<td>" + Экр(СтТов.Номенклатура) + "</td>"
				+ "<td class=""num"">" + Формат(СтТов.Количество, "ЧДЦ=3") + "</td>"
				+ "<td class=""num"">" + ?(СтТов.Цена <> 0, Формат(СтТов.Цена, "ЧДЦ=2"), "&mdash;") + "</td>"
				+ "<td class=""num"">" + ?(СтТов.Сумма <> 0, Формат(СтТов.Сумма, "ЧДЦ=2"), "&mdash;") + "</td>"
				+ "<td>" + Экр(СтТов.СтавкаНДС) + "</td>"
				+ "<td class=""num"">" + ?(СтТов.СуммаНДС <> 0, Формат(СтТов.СуммаНДС, "ЧДЦ=2"), "&mdash;") + "</td></tr>";
		КонецЦикла;
		HTML = HTML + "<tr class=""total-row""><td colspan=""4"">ИТОГО</td>"
			+ "<td class=""num"">" + Формат(ИтогоСумма, "ЧДЦ=2") + "</td><td></td>"
			+ "<td class=""num"">" + Формат(ИтогоНДС, "ЧДЦ=2") + "</td></tr></table>";
	КонецЕсли;

	// --- ОПЛАТЫ (если ОРП) ---
	Если СтрокиОплат.Количество() > 0 Тогда
		HTML = HTML + "<div class=""section"">Оплаты</div><table>"
			+ "<tr><th>Вид оплаты</th><th>Сумма</th></tr>";
		ИтогоОпл = 0;
		Для Каждого СтОпл Из СтрокиОплат Цикл
			ИтогоОпл = ИтогоОпл + СтОпл.Сумма;
			HTML = HTML + "<tr><td>" + Экр(СтОпл.ВидОплаты) + "</td>"
				+ "<td class=""num"">" + Формат(СтОпл.Сумма, "ЧДЦ=2") + "</td></tr>";
		КонецЦикла;
		HTML = HTML + "<tr class=""total-row""><td>ИТОГО</td>"
			+ "<td class=""num"">" + Формат(ИтогоОпл, "ЧДЦ=2") + "</td></tr></table>";
	КонецЕсли;

	// --- ПРОВОДКИ основного документа ---
	HTML = HTML + "<div class=""section"">Проводки (фактические из БД)</div>";
	Если СтрокиПроводок.Количество() > 0 Тогда
		HTML = HTML + "<table><tr><th style=""width:120px"">Дт</th><th style=""width:120px"">Кт</th><th>Сумма</th></tr>";
		Для Каждого СтП Из СтрокиПроводок Цикл
			HTML = HTML + "<tr><td class=""provodka"">" + Экр(СтП.Дт) + "</td>"
				+ "<td class=""provodka"">" + Экр(СтП.Кт) + "</td>"
				+ "<td class=""num"">" + Формат(СтП.Сумма, "ЧДЦ=2") + "</td></tr>";
		КонецЦикла;
		HTML = HTML + "<tr class=""total-row""><td colspan=""2"">ИТОГО оборотов</td>"
			+ "<td class=""num"">" + Формат(ИтогоПроводки, "ЧДЦ=2") + "</td></tr></table>";
	ИначеЕсли ЭтоИнвентаризация Тогда
		HTML = HTML + "<div class=""doc-row disabled""><div class=""doc-name"">"
			+ "Инвентаризация не формирует проводок — фиксирует фактические остатки."
			+ "</div></div>";
	ИначеЕсли Не ПроведёнОсн Тогда
		HTML = HTML + "<div class=""doc-row disabled""><div class=""doc-name"">"
			+ "Документ не проведён — проводки появятся после проведения бухгалтером."
			+ "</div></div>";
	Иначе
		HTML = HTML + "<div class=""doc-row disabled""><div class=""doc-name"">Проводок нет.</div></div>";
	КонецЕсли;

	HTML = HTML + "</body></html>";
	Возврат HTML;

КонецФункции

// Читаемое имя типа документа БП по имени метаданных (для карточки смены из БД).
//
Функция _ДетБД_ИмяТипаДок(ИмяМД)
	Если ИмяМД = "ОтчетОРозничныхПродажах"      Тогда Возврат "ОРП (продажа)"; КонецЕсли;
	Если ИмяМД = "ПоступлениеТоваровУслуг"      Тогда Возврат "Поступление"; КонецЕсли;
	Если ИмяМД = "ИнвентаризацияТоваровНаСкладе" Тогда Возврат "Инвентаризация"; КонецЕсли;
	Если ИмяМД = "ОприходованиеТоваров"         Тогда Возврат "Оприходование"; КонецЕсли;
	Если ИмяМД = "СписаниеТоваров"              Тогда Возврат "Списание"; КонецЕсли;
	Если ИмяМД = "ВозвратТоваровПоставщику"     Тогда Возврат "Возврат поставщику"; КонецЕсли;
	Если ИмяМД = "ПеремещениеТоваров"           Тогда Возврат "Перемещение"; КонецЕсли;
	Если ИмяМД = "ОтчетПроизводстваЗаСмену"     Тогда Возврат "Выпуск продукции"; КонецЕсли;
	Возврат ?(ЗначениеЗаполнено(ИмяМД), ИмяМД, "Документ");
КонецФункции

// Код вида документа для data-action="open|<Вид>|<UUID>" (см. ФормаДетали).
//
Функция _ДетБД_ВидДляОткрытия(ИмяМД)
	Если ИмяМД = "ОтчетОРозничныхПродажах"      Тогда Возврат "ОРП"; КонецЕсли;
	Если ИмяМД = "ПоступлениеТоваровУслуг"      Тогда Возврат "ПТУ"; КонецЕсли;
	Если ИмяМД = "ИнвентаризацияТоваровНаСкладе" Тогда Возврат "Инвентаризация"; КонецЕсли;
	Если ИмяМД = "ОприходованиеТоваров"         Тогда Возврат "Оприходование"; КонецЕсли;
	Если ИмяМД = "СписаниеТоваров"              Тогда Возврат "Списание"; КонецЕсли;
	Если ИмяМД = "ВозвратТоваровПоставщику"     Тогда Возврат "ВозвратПоставщику"; КонецЕсли;
	Если ИмяМД = "ПеремещениеТоваров"           Тогда Возврат "Перемещение"; КонецЕсли;
	Возврат "";
КонецФункции

// Цветной бейдж типа документа (короткий) для колонки «Док» в таблице товаров.
//
Функция _ДетБД_БейджТипа(ИмяМД)
	Если ИмяМД = "ОтчетОРозничныхПродажах" Тогда
		Возврат "<span class=""route-tag route-retail"">ОРП</span>";
	ИначеЕсли ИмяМД = "ПоступлениеТоваровУслуг" Тогда
		Возврат "<span class=""route-tag route-transfer"">Пост</span>";
	ИначеЕсли ИмяМД = "ИнвентаризацияТоваровНаСкладе" Тогда
		Возврат "<span class=""route-tag"" style=""background:#e2e3e5;color:#41464b"">Инв</span>";
	ИначеЕсли ИмяМД = "ОприходованиеТоваров" Тогда
		Возврат "<span class=""route-tag"" style=""background:#e7d6ff;color:#5b21b6"">Опр</span>";
	ИначеЕсли ИмяМД = "СписаниеТоваров" Тогда
		Возврат "<span class=""route-tag"" style=""background:#f8d7da;color:#842029"">Спис</span>";
	ИначеЕсли ИмяМД = "ПеремещениеТоваров" Тогда
		Возврат "<span class=""route-tag"" style=""background:#cff4fc;color:#055160"">Перем</span>";
	Иначе
		Возврат "<span class=""route-tag"" style=""background:#e2e3e5;color:#41464b"">&mdash;</span>";
	КонецЕсли;
КонецФункции

// Экранирование HTML-спецсимволов (защита от XSS в HTMLПоле).
//
Функция Экр(Текст)
	Рез = СтрЗаменить(Строка(Текст), "&", "&amp;");
	Рез = СтрЗаменить(Рез, "<", "&lt;");
	Рез = СтрЗаменить(Рез, ">", "&gt;");
	Рез = СтрЗаменить(Рез, """", "&quot;");
	Возврат Рез;
КонецФункции

// Форматирование суммы (1 234 567 → "1.2М").
//
Функция ФорматСуммы(Сумма)
	Если Сумма >= 1000000 Тогда
		Возврат Формат(Окр(Сумма / 1000000, 1), "ЧДЦ=1") + "М";
	ИначеЕсли Сумма >= 1000 Тогда
		Возврат Формат(Окр(Сумма / 1000, 0), "ЧДЦ=0") + "к";
	Иначе
		Возврат Формат(Сумма, "ЧДЦ=0");
	КонецЕсли;
КонецФункции

// ============================================================================
//  ДЕТАЛИ ПАКЕТА СМЕНЫ (Сопутка/Общепит) — HTML
// ============================================================================
//
// Источник: путь к JSON-файлу пакета. Строит наглядную HTML-страницу с
// KPI-полосой, картой смены, сверкой ЦБ↔БП, ошибками и карточками документов.
//
Функция СформироватьДеталиПакетаСмены(ПутьФайла, Заголовок) Экспорт

	Чтение = TL_HTTPКлиентЦБ.ПрочитатьПакетИзФайла(ПутьФайла);
	Если ЗначениеЗаполнено(Чтение.Ошибка) Или Чтение.Пакет = Неопределено Тогда
		Возврат _ДетПак_ШаблонОшибка("Не удалось прочитать пакет: " + Чтение.Ошибка);
	КонецЕсли;
	Пакет = Чтение.Пакет;

	ИдПакета  = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Пакет, "ИдентификаторПакета", "");
	Версия    = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Пакет, "ВерсияФормата", "1");

	Смена = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Пакет, "Смена", Неопределено);
	КодАЗС = 0; НомерСмены = ""; Открытие = '00010101'; Закрытие = '00010101';
	Оператор = ""; Касса = ""; ОСЭНомер = "";
	СкладUUID = ""; ОрганизацияUUID = "";
	Если ТипЗнч(Смена) = Тип("Соответствие") Тогда
		Попытка КодАЗС = Число(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Смена, "КодАЗС", 0)); Исключение КонецПопытки;
		НомерСмены = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Смена, "НомерСмены", "")));
		Открытие   = _ДетПак_ПарсДата(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Смена, "Открытие", ""));
		Закрытие   = _ДетПак_ПарсДата(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Смена, "Закрытие", ""));
		Оператор   = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Смена, "Оператор", "")));
		Касса      = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Смена, "Касса", "")));
		ОСЭНомер   = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Смена, "ОСЭНомер", "")));
		СкладUUID  = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Смена, "СкладUUID", "")));
		ОрганизацияUUID = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Смена, "ОрганизацияUUID", "")));
	КонецЕсли;

	ДокМассив = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Пакет, "Документы", Новый Массив);
	НСИМассив = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Пакет, "НСИ", Новый Массив);

	Счёт = _ДетПак_Счётчики();
	СуммаПакета = 0;
	Если ТипЗнч(ДокМассив) = Тип("Массив") Или ТипЗнч(ДокМассив) = Тип("ФиксированныйМассив") Тогда
		Для Каждого Д Из ДокМассив Цикл
			ТипД = СокрЛП(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Д, "Тип", ""));
			Если Счёт.Свойство(ТипД) Тогда Счёт[ТипД] = Счёт[ТипД] + 1; КонецЕсли;
			СуммаПакета = СуммаПакета + TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Д, "СуммаДокумента", 0);
		КонецЦикла;
	КонецЕсли;

	НСИКэш = Новый Соответствие;
	Если ТипЗнч(НСИМассив) = Тип("Массив") Или ТипЗнч(НСИМассив) = Тип("ФиксированныйМассив") Тогда
		Для Каждого Эл Из НСИМассив Цикл
			UUID = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Эл, "ИсточникUUID", "")));
			Имя  = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Эл, "Наименование", "")));
			Если ЗначениеЗаполнено(UUID) Тогда НСИКэш.Вставить(UUID, Имя); КонецЕсли;
		КонецЦикла;
	КонецЕсли;

	СопостBP     = _ДетПак_ЗагрузитьСопоставления(ИдПакета);
	СверкаЗапись = _ДетПак_НайтиСверку(ИдПакета);
	СписокОшибок = _ДетПак_НайтиОшибки(ИдПакета);

	РежимРаздельный = TL_Настройки.РежимРаздельногоУчёта();

	HTML = _ДетПак_HTMLШаблон();
	HTML = HTML + _ДетПак_HTMLШапка(Заголовок, КодАЗС, НомерСмены, Открытие, Закрытие, ОСЭНомер, Версия, СуммаПакета, Счёт);
	HTML = HTML + _ДетПак_HTMLКарточкаСмены(КодАЗС, НомерСмены, Открытие, Закрытие, Оператор, Касса, ОСЭНомер, ОрганизацияUUID, СкладUUID, НСИКэш);
	HTML = HTML + _ДетПак_HTMLСверка(СверкаЗапись, СуммаПакета);
	HTML = HTML + _ДетПак_HTMLОшибки(СписокОшибок);
	HTML = HTML + _ДетПак_HTMLДокументы(ДокМассив, НСИКэш, СопостBP);
	HTML = HTML + _ДетПак_HTMLЛегендаПроводок(Счёт, РежимРаздельный);
	HTML = HTML + "</div></body></html>";

	Возврат HTML;

КонецФункции

// Легенда: что приёмник создаст в БП ГИГ при проведении пакета.
// Для каждого kind, который присутствует в пакете (Счёт.<kind> > 0):
//   - тип создаваемого документа БП
//   - типовые бухгалтерские проводки
// Информационные полоски: модель учёта общепита (A/B) + раздельный учёт сопутки/общепита.
//
Функция _ДетПак_HTMLЛегендаПроводок(Счёт, РежимРаздельный = Ложь)

	МодельB = TL_СопуткаСервис.ЭтоМодельОбщепитаB();

	HTML = "<div style='background:#ffffff;border:1px solid #d8dee7;border-radius:8px;padding:18px;margin-top:14px;'>"
		+ "<div style='font-size:13px;font-weight:700;color:#1a1f36;letter-spacing:0.4px;text-transform:uppercase;margin-bottom:12px;'>"
		+ "При проведении пакета будут созданы документы:</div>";

	// Модель учёта общепита (главная информационная полоска)
	Если МодельB Тогда
		HTML = HTML + "<div style='margin-bottom:8px;padding:8px 12px;background:#ecfdf5;border-left:4px solid #059669;font-size:11px;color:#065f46;'>"
			+ "<b>Модель учёта общепита: B</b> — себестоимость блюд списывается прямо на 90.02.1 "
			+ "по ингредиентам ТТК при продаже. <b>ОПЗС не создаётся</b> (kind <code>production_release</code> пропускается). "
			+ "Спецификации блюд приходят отдельным kind <code>recipe</code> из ЦБ ЭЛСИ.АЗК."
			+ "</div>";
	Иначе
		HTML = HTML + "<div style='margin-bottom:8px;padding:8px 12px;background:#fef3c7;border-left:4px solid #d97706;font-size:11px;color:#92400e;'>"
			+ "<b>Модель учёта общепита: A (legacy НЛ)</b> — выпуск через 44.01, ОПЗС создаётся, "
			+ "себестоимость блюд формируется при выпуске."
			+ "</div>";
	КонецЕсли;

	// Раздельный учёт сопутки/общепита (вспомогательная полоска)
	Если РежимРаздельный Тогда
		HTML = HTML + "<div style='margin-bottom:12px;padding:8px 12px;background:#eff6ff;border-left:4px solid #3b82f6;font-size:11px;color:#1e3a8a;'>"
			+ "<b>Раздельный учёт ВКЛ</b> — retail_sale_sidegoods будет разделён на два ОРП: "
			+ "сопутка отдельно от готовой продукции общепита. "
			+ "<i>(на текущем этапе разделение не реализовано — фактически создаётся один сводный ОРП).</i>"
			+ "</div>";
	Иначе
		HTML = HTML + "<div style='margin-bottom:12px;padding:8px 12px;background:#f8fafc;border-left:4px solid #94a3b8;font-size:11px;color:#475569;'>"
			+ "Режим: <b>один сводный ОРП на смену</b> (как у Норд-Лайн). "
			+ "Сопутка и готовая продукция общепита в одном документе."
			+ "</div>";
	КонецЕсли;

	ЕстьСтроки = Ложь;
	HTML = HTML + "<table border='0' cellpadding='6' cellspacing='0' style='border-collapse:collapse;width:100%;font-size:12px;'>"
		+ "<tr style='background:#eef1fb;color:#1a1f36;'>"
		+ "<th style='border:1px solid #d8dee7;text-align:left;padding:8px 10px;width:18%;'>Kind</th>"
		+ "<th style='border:1px solid #d8dee7;text-align:left;padding:8px 10px;width:28%;'>Документ в БП ГИГ</th>"
		+ "<th style='border:1px solid #d8dee7;text-align:left;padding:8px 10px;width:34%;'>Проводки</th>"
		+ "<th style='border:1px solid #d8dee7;text-align:right;padding:8px 10px;width:6%;'>Шт.</th>"
		+ "<th style='border:1px solid #d8dee7;text-align:left;padding:8px 10px;width:14%;'>Примечание</th>"
		+ "</tr>";

	HTML = HTML + _ДетПак_ЛегендаСтрока("purchase", "Поступление товаров и услуг",
		"Дт 41.02 Кт 60.01 (товар)<br>Дт 19.03 Кт 60.01 (НДС)",
		Счёт.purchase, "", Счёт);

	// retail_sale_sidegoods — при модели B добавляем строку списания ингредиентов блюд
	Если МодельB Тогда
		ПроводкиРозн = "Дт 62.Р Кт 90.01.1 (выручка)<br>"
			+ "Дт 90.02.1 Кт 41.02 (с/с сопутки)<br>"
			+ "<b style='color:#059669;'>Дт 90.02.1 Кт 41.02 (ингредиенты блюд по ТТК)</b><br>"
			+ "Дт 90.03 Кт 68.02 (НДС)<br>"
			+ "Дт 50.01/57.03 Кт 62.Р (оплаты)";
		ПримРозн = "При наличии блюд — раскладка по ТТК";
	Иначе
		ПроводкиРозн = "Дт 62.Р Кт 90.01.1 (выручка)<br>"
			+ "Дт 90.02.1 Кт 41.02 (с/с)<br>"
			+ "Дт 90.03 Кт 68.02 (НДС)<br>"
			+ "Дт 50.01/57.03 Кт 62.Р (оплаты)";
		ПримРозн = "";
	КонецЕсли;
	HTML = HTML + _ДетПак_ЛегендаСтрока("retail_sale_sidegoods", "Отчёт о розничных продажах",
		ПроводкиРозн, Счёт.retail_sale_sidegoods, ПримРозн, Счёт);

	// production_release — при модели B пропускается, при A через 44.01
	Если МодельB Тогда
		ПроводкиПрод = "<i style='color:#94a3b8;'>не создаётся</i>";
		ИмяДокПрод   = "<i style='color:#94a3b8;'>пропускается (модель B)</i>";
		ПримПрод     = "с/с уже в retail_sale_sidegoods";
	Иначе
		ПроводкиПрод = "Дт 41.02 Кт 44.01 (выпуск)<br>Дт 44.01 Кт 41.02 (списание сырья)";
		ИмяДокПрод   = "Отчёт производства за смену";
		ПримПрод     = "";
	КонецЕсли;
	HTML = HTML + _ДетПак_ЛегендаСтрока("production_release", ИмяДокПрод,
		ПроводкиПрод, Счёт.production_release, ПримПрод, Счёт);

	HTML = HTML + _ДетПак_ЛегендаСтрока("recipe", "Спецификация номенклатуры",
		"Без проводок (НСИ)",
		Счёт.recipe, "ТТК блюда из ЦБ", Счёт);

	HTML = HTML + _ДетПак_ЛегендаСтрока("return_purchase", "Корректировка поступления",
		"Дт 41.02 Кт 60.01 СТОРНО<br>Дт 19.03 Кт 60.01 СТОРНО",
		Счёт.return_purchase, "", Счёт);

	HTML = HTML + _ДетПак_ЛегендаСтрока("inventory", "Инвентаризация товаров на складе",
		"Без проводок (носитель факта)",
		Счёт.inventory, "", Счёт);

	HTML = HTML + _ДетПак_ЛегендаСтрока("gain", "Оприходование товаров",
		"Дт 41.02 Кт 91.01 (излишки)",
		Счёт.gain, "", Счёт);

	HTML = HTML + _ДетПак_ЛегендаСтрока("writeoff", "Списание товаров",
		"Дт 94 Кт 41.02 (недостача)<br><i style='color:#94a3b8;'>или Дт 90.02.1 Кт 41.02 при канале writeoff_fuel</i>",
		Счёт.writeoff, "", Счёт);

	HTML = HTML + _ДетПак_ЛегендаСтрока("transfer", "Перемещение товаров",
		"Дт 41.02 Кт 41.02 (между складами)",
		Счёт.transfer, "", Счёт);

	HTML = HTML + "</table>";

	HTML = HTML + "</div>";
	Возврат HTML;

КонецФункции

// Одна строка таблицы легенды.
//
Функция _ДетПак_ЛегендаСтрока(Kind, Документ, Проводки, Количество, Примечание, Счёт)
	// Если этот kind в пакете отсутствует — серая строка-заглушка для полноты картины
	ЦветТекста = ?(Количество > 0, "#1a1f36", "#94a3b8");
	HTML = "<tr>"
		+ "<td style='border:1px solid #e5e9f2;padding:6px 10px;color:" + ЦветТекста + ";'><code>" + Kind + "</code></td>"
		+ "<td style='border:1px solid #e5e9f2;padding:6px 10px;color:" + ЦветТекста + ";'>" + Документ + "</td>"
		+ "<td style='border:1px solid #e5e9f2;padding:6px 10px;color:" + ЦветТекста + ";font-size:11px;line-height:1.55;'>" + Проводки + "</td>"
		+ "<td style='border:1px solid #e5e9f2;padding:6px 10px;text-align:right;color:" + ЦветТекста + ";font-weight:" + ?(Количество > 0, "600", "400") + ";'>"
		+ ?(Количество > 0, Формат(Количество, "ЧГ=' '; ЧН=0"), "—") + "</td>"
		+ "<td style='border:1px solid #e5e9f2;padding:6px 10px;font-size:11px;'>" + Примечание + "</td>"
		+ "</tr>";
	Возврат HTML;
КонецФункции

Функция _ДетПак_Счётчики()
	Стр = Новый Структура;
	Стр.Вставить("purchase", 0);
	Стр.Вставить("retail_sale_sidegoods", 0);
	Стр.Вставить("production_release", 0);
	Стр.Вставить("recipe", 0);
	Стр.Вставить("return_purchase", 0);
	Стр.Вставить("inventory", 0);
	Стр.Вставить("gain", 0);
	Стр.Вставить("writeoff", 0);
	Стр.Вставить("transfer", 0);
	Возврат Стр;
КонецФункции

Функция _ДетПак_ПарсДата(Стр)
	Если Не ЗначениеЗаполнено(Стр) Тогда Возврат '00010101'; КонецЕсли;
	Попытка
		Возврат XMLЗначение(Тип("Дата"), Строка(Стр));
	Исключение
		Возврат '00010101';
	КонецПопытки;
КонецФункции

// Палитра типа документа — refined-схема для бухгалтерского UI:
//   Граница — насыщенный 600-й оттенок (border-left документа, маркер цвета)
//   Фон     — мягкий 50-й оттенок (бейдж, читается с тёмным текстом)
//   Текст   — глубокий 700-800-й (контраст AA на фоне 50-го >= 7:1)
// Все цвета из Tailwind-палитры — выверенный контраст без подбора.
//
Функция _ДетПак_ПалитраТипа(Тип)
	П = Новый Структура("Граница, Фон, Текст");
	Если Тип = "purchase" Тогда
		П.Граница = "#2563EB"; П.Фон = "#EFF6FF"; П.Текст = "#1E40AF";
	ИначеЕсли Тип = "retail_sale_sidegoods" Тогда
		П.Граница = "#059669"; П.Фон = "#ECFDF5"; П.Текст = "#065F46";
	ИначеЕсли Тип = "production_release" Тогда
		П.Граница = "#D97706"; П.Фон = "#FEF3C7"; П.Текст = "#92400E";
	ИначеЕсли Тип = "recipe" Тогда
		П.Граница = "#10B981"; П.Фон = "#D1FAE5"; П.Текст = "#047857";
	ИначеЕсли Тип = "return_purchase" Тогда
		П.Граница = "#E11D48"; П.Фон = "#FFE4E6"; П.Текст = "#9F1239";
	ИначеЕсли Тип = "inventory" Тогда
		П.Граница = "#64748B"; П.Фон = "#F1F5F9"; П.Текст = "#334155";
	ИначеЕсли Тип = "gain" Тогда
		П.Граница = "#7C3AED"; П.Фон = "#F3E8FF"; П.Текст = "#5B21B6";
	ИначеЕсли Тип = "writeoff" Тогда
		П.Граница = "#DC2626"; П.Фон = "#FEE2E2"; П.Текст = "#991B1B";
	ИначеЕсли Тип = "transfer" Тогда
		П.Граница = "#0891B2"; П.Фон = "#CFFAFE"; П.Текст = "#155E75";
	Иначе
		П.Граница = "#64748B"; П.Фон = "#F1F5F9"; П.Текст = "#334155";
	КонецЕсли;
	Возврат П;
КонецФункции

// Алиас (Граница) для обратной совместимости с местами, где используется одиночный цвет.
Функция _ДетПак_ЦветТипа(Тип)
	Возврат _ДетПак_ПалитраТипа(Тип).Граница;
КонецФункции

Функция _ДетПак_ИмяТипа(Тип)
	Если Тип = "purchase"             Тогда Возврат "Поступление"; КонецЕсли;
	Если Тип = "retail_sale_sidegoods" Тогда Возврат "Розничные продажи"; КонецЕсли;
	Если Тип = "production_release"    Тогда Возврат "Выпуск продукции"; КонецЕсли;
	Если Тип = "recipe"                Тогда Возврат "Спецификация (ТТК)"; КонецЕсли;
	Если Тип = "return_purchase"       Тогда Возврат "Возврат поставщику"; КонецЕсли;
	Если Тип = "inventory"             Тогда Возврат "Инвентаризация"; КонецЕсли;
	Если Тип = "gain"                  Тогда Возврат "Оприходование"; КонецЕсли;
	Если Тип = "writeoff"              Тогда Возврат "Списание"; КонецЕсли;
	Если Тип = "transfer"              Тогда Возврат "Перемещение"; КонецЕсли;
	Возврат Тип;
КонецФункции

Функция _ДетПак_ФорматСуммы(Сумма)
	Возврат Формат(Сумма, "ЧДЦ=2; ЧРГ=' '") + " ₽";
КонецФункции

Функция _ДетПак_HTMLШаблон()
	Стиль = "body{margin:0;padding:0;font-family:-apple-system,'Segoe UI',sans-serif;font-size:13px;background:#F0F4F8;color:#1F2937;}"
		+ ".wrap{padding:12px 16px;}"
		+ "h1{font-size:18px;margin:0 0 4px 0;color:#1F2937;}"
		+ ".sub{font-size:11px;color:#6B7280;margin-bottom:12px;}"
		+ ".kpi{display:grid;grid-template-columns:repeat(9,1fr);gap:6px;margin-bottom:12px;}"
		+ ".kpi-card{background:#fff;border:1px solid #E2E8F0;border-radius:6px;padding:8px;text-align:center;}"
		+ ".kpi-card .val{font-size:18px;font-weight:600;}"
		+ ".kpi-card .lbl{font-size:10px;color:#6B7280;text-transform:uppercase;letter-spacing:0.4px;}"
		+ ".section{background:#fff;border:1px solid #E2E8F0;border-radius:6px;padding:10px 12px;margin-bottom:10px;}"
		+ ".section h2{font-size:13px;margin:0 0 8px 0;color:#374151;text-transform:uppercase;letter-spacing:0.5px;font-weight:600;}"
		+ ".meta{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;font-size:12px;}"
		+ ".meta .lbl{color:#6B7280;font-size:10px;text-transform:uppercase;letter-spacing:0.4px;}"
		+ ".meta .val{font-weight:500;}"
		+ ".doc{background:#fff;border:1px solid #E2E8F0;border-left:4px solid;border-radius:6px;padding:10px 12px;margin-bottom:8px;}"
		+ ".doc-hdr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px;}"
		+ ".doc-hdr .left{font-weight:600;font-size:14px;}"
		+ ".doc-hdr .right{text-align:right;font-size:13px;}"
		+ ".badge{display:inline-block;font-size:11px;padding:3px 9px;border-radius:4px;text-transform:uppercase;font-weight:700;letter-spacing:0.3px;line-height:1.3;text-decoration:none;white-space:nowrap;border:1px solid transparent;}"
		+ ".doc-meta{font-size:11px;color:#6B7280;margin-bottom:6px;}"
		+ ".doc-meta span{margin-right:12px;}"
		+ ".tbl{width:100%;border-collapse:collapse;font-size:11px;margin-top:6px;}"
		+ ".tbl th{background:#F8FAFC;color:#475569;font-weight:600;text-align:left;padding:4px 6px;border-bottom:1px solid #E2E8F0;}"
		+ ".tbl td{padding:3px 6px;border-bottom:1px solid #F1F5F9;}"
		+ ".tbl td.num{text-align:right;font-variant-numeric:tabular-nums;}"
		+ ".tbl tr:hover{background:#F8FAFC;}"
		+ ".warn{background:#FEF3C7;color:#92400E;padding:8px 10px;border-radius:4px;border-left:4px solid #EAB308;margin-bottom:10px;}"
		+ ".err{background:#FEE2E2;color:#991B1B;padding:8px 10px;border-radius:4px;border-left:4px solid #EF4444;margin-bottom:10px;}"
		+ ".ok-pill{display:inline-block;background:#DCFCE7;color:#166534;padding:2px 9px;border-radius:10px;font-size:10.5px;font-weight:700;border:1px solid #BBF7D0;letter-spacing:0.2px;}"
		+ ".ko-pill{display:inline-block;background:#FEE2E2;color:#991B1B;padding:2px 9px;border-radius:10px;font-size:10.5px;font-weight:700;border:1px solid #FECACA;letter-spacing:0.2px;}"
		+ "details,summary{outline:none;}"
		+ "summary{text-decoration:none;}"
		+ "details.doc{background:#fff;border:1px solid #E2E8F0;border-left:5px solid;border-radius:6px;padding:0;margin-bottom:7px;overflow:hidden;transition:box-shadow .15s,border-color .15s;}"
		+ "details.doc:hover{box-shadow:0 1px 3px rgba(15,23,42,0.08);border-color:#CBD5E1;}"
		+ "details.doc[open]{box-shadow:0 2px 6px rgba(15,23,42,0.06);}"
		+ "details.doc>summary{cursor:pointer;padding:10px 14px;list-style:none;display:flex;justify-content:space-between;align-items:center;gap:10px;user-select:none;text-decoration:none;}"
		+ "details.doc>summary::-webkit-details-marker{display:none;}"
		+ "details.doc>summary::before{content:'▸';font-size:11px;color:#94A3B8;margin-right:8px;transition:transform .15s;display:inline-block;flex-shrink:0;}"
		+ "details.doc[open]>summary::before{transform:rotate(90deg);color:#475569;}"
		+ "details.doc>summary:hover{background:#F8FAFC;}"
		+ "details.doc>.body{padding:8px 14px 12px 36px;border-top:1px solid #E2E8F0;background:#FAFBFC;}"
		+ "details.docs-list{background:transparent;border:none;padding:0;margin:0;}"
		+ "details.docs-list>summary{cursor:pointer;padding:8px 0;font-size:12px;font-weight:700;color:#1F2937;text-transform:uppercase;letter-spacing:0.6px;list-style:none;text-decoration:none;}"
		+ "details.docs-list>summary::-webkit-details-marker{display:none;}"
		+ "details.docs-list>summary::before{content:'▸';font-size:10px;color:#64748B;margin-right:6px;transition:transform .15s;display:inline-block;}"
		+ "details.docs-list[open]>summary::before{transform:rotate(90deg);}"
		+ "details.tch{margin-top:6px;background:#fff;border:1px solid #E2E8F0;border-radius:4px;overflow:hidden;}"
		+ "details.tch>summary{cursor:pointer;font-size:11px;font-weight:600;color:#475569;padding:6px 10px;list-style:none;user-select:none;text-decoration:none;background:#F8FAFC;}"
		+ "details.tch>summary:hover{background:#F1F5F9;}"
		+ "details.tch>summary::-webkit-details-marker{display:none;}"
		+ "details.tch>summary::before{content:'▸';font-size:9px;color:#94A3B8;margin-right:6px;transition:transform .15s;display:inline-block;}"
		+ "details.tch[open]>summary::before{transform:rotate(90deg);color:#475569;}"
		+ "details.tch>.tbl{margin-top:0;}";
	Возврат "<!DOCTYPE html><html><head><meta charset=""utf-8""><style>" + Стиль + "</style></head><body><div class=""wrap"">";
КонецФункции

Функция _ДетПак_HTMLШапка(Заголовок, КодАЗС, НомерСмены, Открытие, Закрытие, ОСЭНомер, Версия, СуммаПакета, Счёт)
	ДнДата = Формат(Закрытие, "ДФ='dd.MM.yyyy HH:mm'");
	ОткрСтр = Формат(Открытие, "ДФ='dd.MM HH:mm'");
	Заг = "Смена " + НомерСмены + " · АЗС " + Формат(КодАЗС, "ЧГ=0") + " · " + ОткрСтр + " → " + ДнДата;
	Длительность = (Закрытие - Открытие) / 3600;
	СубТитр = "ОСЭ " + ОСЭНомер + " · длительность " + Формат(Длительность, "ЧДЦ=1") + " ч · формат v" + Версия;

	HTML = "<h1>" + Заг + "</h1><div class=""sub"">" + СубТитр + "</div>";
	HTML = HTML + "<div class=""kpi"">";
	HTML = HTML + _ДетПак_KPIКарточка(Счёт.purchase,             "Поступл.",   _ДетПак_ПалитраТипа("purchase").Граница);
	HTML = HTML + _ДетПак_KPIКарточка(Счёт.retail_sale_sidegoods, "Розница",    _ДетПак_ПалитраТипа("retail_sale_sidegoods").Граница);
	HTML = HTML + _ДетПак_KPIКарточка(Счёт.production_release,    "Произв.",    _ДетПак_ПалитраТипа("production_release").Граница);
	HTML = HTML + _ДетПак_KPIКарточка(Счёт.return_purchase,       "Возвр.",     _ДетПак_ПалитраТипа("return_purchase").Граница);
	HTML = HTML + _ДетПак_KPIКарточка(Счёт.inventory,             "Инвент.",    _ДетПак_ПалитраТипа("inventory").Граница);
	HTML = HTML + _ДетПак_KPIКарточка(Счёт.gain,                  "Оприход.",   _ДетПак_ПалитраТипа("gain").Граница);
	HTML = HTML + _ДетПак_KPIКарточка(Счёт.writeoff,              "Списан.",    _ДетПак_ПалитраТипа("writeoff").Граница);
	HTML = HTML + _ДетПак_KPIКарточка(Счёт.transfer,              "Перем.",     _ДетПак_ПалитраТипа("transfer").Граница);

	HTML = HTML + "<div class=""kpi-card"" style=""background:#1F2937;color:#fff;"">"
		+ "<div class=""val"" style=""color:#fff;"">" + Формат(СуммаПакета, "ЧДЦ=0; ЧРГ=' '") + "</div>"
		+ "<div class=""lbl"" style=""color:#9CA3AF;"">Сумма ₽</div></div>";
	HTML = HTML + "</div>";
	Возврат HTML;
КонецФункции

Функция _ДетПак_KPIКарточка(Значение, Метка, Цвет)
	Возврат "<div class=""kpi-card"">"
		+ "<div class=""val"" style=""color:" + Цвет + ";"">" + Формат(Значение, "ЧГ=0") + "</div>"
		+ "<div class=""lbl"">" + Метка + "</div></div>";
КонецФункции

Функция _ДетПак_HTMLКарточкаСмены(КодАЗС, НомерСмены, Открытие, Закрытие, Оператор, Касса, ОСЭНомер, ОргUUID, СклUUID, НСИКэш)
	ОргИмя  = НСИКэш.Получить(ОргUUID);
	СклИмя  = НСИКэш.Получить(СклUUID);
	ОргТекст = ?(ОргИмя <> Неопределено И ЗначениеЗаполнено(ОргИмя), ОргИмя, ОргUUID);
	СклТекст = ?(СклИмя <> Неопределено И ЗначениеЗаполнено(СклИмя), СклИмя, СклUUID);

	HTML = "<div class=""section""><h2>Шапка смены</h2><div class=""meta"">"
		+ _ДетПак_МетаПара("КодАЗС", Формат(КодАЗС, "ЧГ=0"))
		+ _ДетПак_МетаПара("Номер смены", НомерСмены)
		+ _ДетПак_МетаПара("Открытие", Формат(Открытие, "ДФ='dd.MM.yyyy HH:mm:ss'"))
		+ _ДетПак_МетаПара("Закрытие", Формат(Закрытие, "ДФ='dd.MM.yyyy HH:mm:ss'"))
		+ _ДетПак_МетаПара("Оператор", ?(ЗначениеЗаполнено(Оператор), Оператор, "—"))
		+ _ДетПак_МетаПара("Касса", ?(ЗначениеЗаполнено(Касса), Касса, "—"))
		+ _ДетПак_МетаПара("ОСЭ_Номер", ОСЭНомер)
		+ _ДетПак_МетаПара("Организация", ОргТекст)
		+ _ДетПак_МетаПара("Склад", СклТекст);
	HTML = HTML + "</div></div>";
	Возврат HTML;
КонецФункции

Функция _ДетПак_МетаПара(Метка, Значение)
	Возврат "<div><div class=""lbl"">" + Метка + "</div><div class=""val"">" + Значение + "</div></div>";
КонецФункции

Функция _ДетПак_HTMLСверка(СверкаЗапись, СуммаПакетаJSON)
	Если СверкаЗапись = Неопределено Тогда
		Возврат "<div class=""warn"">Пакет ещё не загружен в БП ГИГ (нет записи в TL_СверкаПакета). Нажмите «Загрузить выбранные».</div>";
	КонецЕсли;
	Расх = СверкаЗапись.СуммаЦБ - СверкаЗапись.СуммаБП;
	ЦветРасх = ?(Расх = 0, "#15803D", ?(Расх > 0, "#EAB308", "#EF4444"));
	Статус = Строка(СверкаЗапись.Статус);
	HTML = "<div class=""section""><h2>Сверка ЦБ ↔ БП ГИГ</h2><div class=""meta"">"
		+ _ДетПак_МетаПара("Время загрузки", Формат(СверкаЗапись.ВремяЗагрузки, "ДФ='dd.MM.yyyy HH:mm:ss'"))
		+ _ДетПак_МетаПара("Статус", Статус)
		+ _ДетПак_МетаПара("Создано документов", Формат(СверкаЗапись.СтрокДокументов, "ЧГ=0"))
		+ _ДетПак_МетаПара("Сумма ЦБ", _ДетПак_ФорматСуммы(СверкаЗапись.СуммаЦБ))
		+ _ДетПак_МетаПара("Сумма БП", _ДетПак_ФорматСуммы(СверкаЗапись.СуммаБП))
		+ "<div><div class=""lbl"">Расхождение</div><div class=""val"" style=""color:" + ЦветРасх + ";font-weight:600;"">"
		+ _ДетПак_ФорматСуммы(Расх) + "</div></div>"
		+ "</div></div>";
	Возврат HTML;
КонецФункции

Функция _ДетПак_HTMLОшибки(СписокОшибок)
	Если СписокОшибок.Количество() = 0 Тогда
		Возврат "";
	КонецЕсли;
	HTML = "<div class=""section""><h2 style=""color:#991B1B;"">Ошибки загрузки (" + СписокОшибок.Количество() + ")</h2>";
	Для Каждого Стр Из СписокОшибок Цикл
		HTML = HTML + "<div class=""err"">"
			+ "<b>" + Стр.Сообщение + "</b>"
			+ ?(ЗначениеЗаполнено(Стр.КодОшибки), " · код: " + Стр.КодОшибки, "")
			+ ?(ЗначениеЗаполнено(Стр.ТипОбъекта), " · объект: " + Стр.ТипОбъекта, "")
			+ ?(ЗначениеЗаполнено(Стр.ЗначениеИсточника), " · UUID: " + Стр.ЗначениеИсточника, "")
			+ "</div>";
	КонецЦикла;
	HTML = HTML + "</div>";
	Возврат HTML;
КонецФункции

Функция _ДетПак_HTMLДокументы(ДокМассив, НСИКэш, СопостBP)
	Если ТипЗнч(ДокМассив) <> Тип("Массив") И ТипЗнч(ДокМассив) <> Тип("ФиксированныйМассив") Тогда
		Возврат "";
	КонецЕсли;
	// Внешний блок с документами раскрыт по умолчанию (виден список),
	// каждый документ внутри — свёрнут (раскрытие по клику).
	HTML = "<div class=""section""><details class=""docs-list"" open>"
		+ "<summary>Документы пакета (" + ДокМассив.Количество() + ") — кликните для свёртки списка</summary>"
		+ "<div style=""margin-top:6px;"">";
	НомерДок = 0;
	Для Каждого Д Из ДокМассив Цикл
		НомерДок = НомерДок + 1;
		HTML = HTML + _ДетПак_ОдинДок(Д, НомерДок, НСИКэш, СопостBP);
	КонецЦикла;
	HTML = HTML + "</div></details></div>";
	Возврат HTML;
КонецФункции

Функция _ДетПак_ОдинДок(Док, НомерДок, НСИКэш, СопостBP)
	Тип    = СокрЛП(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "Тип", ""));
	UUID   = СокрЛП(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "ИсточникUUID", ""));
	Номер  = СокрЛП(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "Номер", ""));
	Дата   = _ДетПак_ПарсДата(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "Дата", ""));
	Сумма  = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "СуммаДокумента", 0);
	П      = _ДетПак_ПалитраТипа(Тип);
	ИмяТип = _ДетПак_ИмяТипа(Тип);

	Ссылка = СопостBP.Получить(UUID);
	СтатусСтр = "";
	Если Ссылка <> Неопределено Тогда
		Попытка
			Проведен = Ссылка.Проведен;
			СтатусСтр = "<span class=""ok-pill"">" + ?(Проведен, "✓ Проведён", "● Записан") + "</span> "
				+ "<span style=""font-size:11px;color:#64748B;font-weight:400;"">" + Строка(Ссылка) + "</span>";
		Исключение
			СтатусСтр = "<span class=""ok-pill"">создан</span>";
		КонецПопытки;
	Иначе
		СтатусСтр = "<span class=""ko-pill"">не создан</span>";
	КонецЕсли;

	// Каждый документ — свёрнутый по умолчанию <details>. В заголовке (summary) —
	// тип, номер, статус, сумма (видно при свёрнутом состоянии). При раскрытии —
	// полная мета и ТЧ.
	HTML = "<details class=""doc"" style=""border-left-color:" + П.Граница + ";"">";
	HTML = HTML + "<summary>"
		+ "<div class=""left"" style=""display:flex;align-items:center;gap:10px;flex:1;min-width:0;"">"
		+ "<span class=""badge"" style=""background:" + П.Фон + ";color:" + П.Текст
		+ ";border-color:" + П.Граница + ";"">" + ИмяТип + "</span>"
		+ "<span style=""color:#94A3B8;font-weight:400;"">#" + Формат(НомерДок, "ЧГ=0") + "</span>"
		+ ?(ЗначениеЗаполнено(Номер),
			"<span style=""font-family:'SF Mono','Consolas',monospace;font-size:12px;color:#1E293B;font-weight:600;"">" + Номер + "</span>",
			"")
		+ "<span style=""font-size:11px;color:#64748B;font-weight:400;"">" + Формат(Дата, "ДФ='dd.MM HH:mm:ss'") + "</span>"
		+ СтатусСтр + "</div>"
		+ "<div class=""right"" style=""font-weight:700;font-size:14px;color:#0F172A;text-align:right;flex-shrink:0;font-variant-numeric:tabular-nums;"">"
		+ _ДетПак_ФорматСуммы(Сумма) + "</div>"
		+ "</summary>";
	HTML = HTML + "<div class=""body"">";
	HTML = HTML + "<div class=""doc-meta"">" + _ДетПак_МетаДок(Док, НСИКэш) + "</div>";
	HTML = HTML + _ДетПак_ТЧТовары(Док, НСИКэш);

	Если Тип = "production_release" Тогда
		HTML = HTML + _ДетПак_ТЧИнгредиенты(Док, НСИКэш);
	КонецЕсли;
	Если Тип = "retail_sale_sidegoods" Тогда
		HTML = HTML + _ДетПак_ТЧОплаты(Док);
	КонецЕсли;

	HTML = HTML + "</div></details>";
	Возврат HTML;
КонецФункции

Функция _ДетПак_МетаДок(Док, НСИКэш)
	Куски = Новый Массив;
	КонтрUUID = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "Контрагент", "")));
	Если ЗначениеЗаполнено(КонтрUUID) Тогда
		Имя = НСИКэш.Получить(КонтрUUID);
		Куски.Добавить("<span><b>Контрагент:</b> " + ?(Имя = Неопределено, КонтрUUID, Имя) + "</span>");
	КонецЕсли;
	СклUUID = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "Склад", "")));
	Если ЗначениеЗаполнено(СклUUID) Тогда
		Имя = НСИКэш.Получить(СклUUID);
		Куски.Добавить("<span><b>Склад:</b> " + ?(Имя = Неопределено, "—", Имя) + "</span>");
	КонецЕсли;
	ОтпрUUID = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "СкладОтправитель", "")));
	ПолучUUID = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "СкладПолучатель", "")));
	Если ЗначениеЗаполнено(ОтпрUUID) И ЗначениеЗаполнено(ПолучUUID) Тогда
		Имя1 = НСИКэш.Получить(ОтпрUUID); Имя2 = НСИКэш.Получить(ПолучUUID);
		Напр = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "Направление", "");
		Куски.Добавить("<span><b>" + ?(Имя1=Неопределено,"?",Имя1) + " → " + ?(Имя2=Неопределено,"?",Имя2)
			+ "</b> · " + Напр + "</span>");
	КонецЕсли;
	Подр = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "Подразделение", "")));
	Если ЗначениеЗаполнено(Подр) Тогда
		Куски.Добавить("<span><b>Подразделение:</b> " + Подр + "</span>");
	КонецЕсли;
	Возврат СтрСоединить(Куски, " ");
КонецФункции

Функция _ДетПак_ТЧТовары(Док, НСИКэш)
	Тип = СокрЛП(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "Тип", ""));
	ИмяТЧ = ?(Тип = "production_release", "ВыпускБлюд", "Товары");
	Массив = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, ИмяТЧ, Новый Массив);
	Если ТипЗнч(Массив) <> Тип("Массив") И ТипЗнч(Массив) <> Тип("ФиксированныйМассив") Тогда
		Возврат "";
	КонецЕсли;
	Если Массив.Количество() = 0 Тогда
		Возврат "<div style=""font-size:11px;color:#6B7280;font-style:italic;"">ТЧ " + ИмяТЧ + " пуста</div>";
	КонецЕсли;

	ЕстьУчетКол = (Тип = "inventory");

	HTML = "<details class=""tch"" open><summary>"
		+ ИмяТЧ + " (" + Массив.Количество() + ")</summary>";
	HTML = HTML + "<table class=""tbl""><thead><tr>"
		+ "<th style=""width:30px;"">#</th>"
		+ "<th>Номенклатура</th>"
		+ "<th style=""text-align:right;width:70px;"">Кол</th>"
		+ ?(ЕстьУчетКол, "<th style=""text-align:right;width:70px;color:#6B7280;"">Уч.кол</th>", "")
		+ ?(ЕстьУчетКол, "<th style=""text-align:right;width:70px;color:#EF4444;"">Δ</th>", "")
		+ "<th style=""width:50px;"">Ед.</th>"
		+ "<th style=""text-align:right;width:80px;"">Цена</th>"
		+ "<th style=""text-align:right;width:90px;"">Сумма</th>"
		+ "<th style=""width:60px;"">НДС</th>"
		+ "<th style=""text-align:right;width:80px;"">∑ НДС</th>"
		+ "</tr></thead><tbody>";

	Для Каждого Стр Из Массив Цикл
		НомСтр = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "НомерСтроки", 0);
		НомUUID = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "Номенклатура", "")));
		НомИмя = НСИКэш.Получить(НомUUID);
		Если НомИмя = Неопределено Тогда НомИмя = "[" + Лев(НомUUID, 8) + "...]"; КонецЕсли;
		Кол  = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "Количество", 0);
		Цена = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "Цена", 0);
		Сум  = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "Сумма", 0);
		НДС  = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "СтавкаНДС", "")));
		СумНДС = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "СуммаНДС", 0);
		Ед   = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "Единица", "")));

		HTML = HTML + "<tr><td>" + Формат(НомСтр, "ЧГ=0") + "</td>"
			+ "<td>" + НомИмя + "</td>"
			+ "<td class=""num"">" + Формат(Кол, "ЧДЦ=3") + "</td>";
		Если ЕстьУчетКол Тогда
			КолУч = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "КоличествоУчет", 0);
			Дельта = Кол - КолУч;
			ЦветД = ?(Дельта = 0, "#6B7280", ?(Дельта > 0, "#16A34A", "#EF4444"));
			HTML = HTML + "<td class=""num"" style=""color:#6B7280;"">" + Формат(КолУч, "ЧДЦ=3") + "</td>"
				+ "<td class=""num"" style=""color:" + ЦветД + ";font-weight:600;"">"
				+ ?(Дельта > 0, "+", "") + Формат(Дельта, "ЧДЦ=3") + "</td>";
		КонецЕсли;
		HTML = HTML + "<td>" + Ед + "</td>"
			+ "<td class=""num"">" + Формат(Цена, "ЧДЦ=2") + "</td>"
			+ "<td class=""num""><b>" + Формат(Сум, "ЧДЦ=2") + "</b></td>"
			+ "<td>" + НДС + "</td>"
			+ "<td class=""num"">" + Формат(СумНДС, "ЧДЦ=2") + "</td></tr>";
	КонецЦикла;
	HTML = HTML + "</tbody></table></details>";
	Возврат HTML;
КонецФункции

Функция _ДетПак_ТЧИнгредиенты(Док, НСИКэш)
	Массив = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "Ингредиенты", Новый Массив);
	Если ТипЗнч(Массив) <> Тип("Массив") И ТипЗнч(Массив) <> Тип("ФиксированныйМассив") Тогда
		Возврат "";
	КонецЕсли;
	Если Массив.Количество() = 0 Тогда Возврат ""; КонецЕсли;

	HTML = "<details class=""tch""><summary>Ингредиенты (" + Массив.Количество() + ")</summary>";
	HTML = HTML + "<table class=""tbl""><thead><tr>"
		+ "<th style=""width:30px;"">#</th>"
		+ "<th>Номенклатура</th>"
		+ "<th>Продукция-источник</th>"
		+ "<th style=""text-align:right;width:90px;"">Кол</th>"
		+ "<th style=""width:50px;"">Ед.</th>"
		+ "</tr></thead><tbody>";
	Для Каждого Стр Из Массив Цикл
		НомСтр = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "НомерСтроки", 0);
		НомUUID = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "Номенклатура", "")));
		НомИмя = НСИКэш.Получить(НомUUID);
		Если НомИмя = Неопределено Тогда НомИмя = "[" + Лев(НомUUID, 8) + "...]"; КонецЕсли;
		ИдПрод = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "ИдентификаторПродукция", "")));
		Кол = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "Количество", 0);
		Ед = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "Единица", "")));
		HTML = HTML + "<tr><td>" + Формат(НомСтр, "ЧГ=0") + "</td>"
			+ "<td>" + НомИмя + "</td>"
			+ "<td style=""font-size:10px;color:#6B7280;"">" + ИдПрод + "</td>"
			+ "<td class=""num"">" + Формат(Кол, "ЧДЦ=3") + "</td>"
			+ "<td>" + Ед + "</td></tr>";
	КонецЦикла;
	HTML = HTML + "</tbody></table></details>";
	Возврат HTML;
КонецФункции

Функция _ДетПак_ТЧОплаты(Док)
	Массив = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Док, "Оплаты", Новый Массив);
	Если ТипЗнч(Массив) <> Тип("Массив") И ТипЗнч(Массив) <> Тип("ФиксированныйМассив") Тогда
		Возврат "";
	КонецЕсли;
	Если Массив.Количество() = 0 Тогда Возврат ""; КонецЕсли;

	HTML = "<details class=""tch""><summary>Оплаты (" + Массив.Количество() + ")</summary>";
	HTML = HTML + "<table class=""tbl""><thead><tr>"
		+ "<th>Вид оплаты</th>"
		+ "<th style=""text-align:right;width:120px;"">Сумма</th>"
		+ "</tr></thead><tbody>";
	Для Каждого Стр Из Массив Цикл
		Вид = СокрЛП(Строка(TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "ВидОплаты", "")));
		Сум = TL_HTTPКлиентЦБ.ПолучитьЗначениеИзJSON(Стр, "Сумма", 0);
		HTML = HTML + "<tr><td>" + Вид + "</td><td class=""num""><b>" + _ДетПак_ФорматСуммы(Сум) + "</b></td></tr>";
	КонецЦикла;
	HTML = HTML + "</tbody></table></details>";
	Возврат HTML;
КонецФункции

Функция _ДетПак_ЗагрузитьСопоставления(ИдПакета)
	Результат = Новый Соответствие;
	Если Не ЗначениеЗаполнено(ИдПакета) Тогда Возврат Результат; КонецЕсли;
	Попытка
		Запрос = Новый Запрос(
			"ВЫБРАТЬ ИсточникUUID, СсылкаОбъект
			|ИЗ РегистрСведений.TL_СоответствиеИсточников
			|ГДЕ ИдентификаторПоследнегоПакета = &Ид");
		Запрос.УстановитьПараметр("Ид", ИдПакета);
		В = Запрос.Выполнить().Выбрать();
		Пока В.Следующий() Цикл
			Результат.Вставить(СокрЛП(В.ИсточникUUID), В.СсылкаОбъект);
		КонецЦикла;
	Исключение
	КонецПопытки;
	Возврат Результат;
КонецФункции

Функция _ДетПак_НайтиСверку(ИдПакета)
	Если Не ЗначениеЗаполнено(ИдПакета) Тогда Возврат Неопределено; КонецЕсли;
	Попытка
		Запрос = Новый Запрос(
			"ВЫБРАТЬ ПЕРВЫЕ 1 ВремяЗагрузки, Статус, СтрокДокументов, СуммаЦБ, СуммаБП
			|ИЗ РегистрСведений.TL_СверкаПакета
			|ГДЕ ИдентификаторПакета = &Ид
			|УПОРЯДОЧИТЬ ПО ВремяЗагрузки УБЫВ");
		Запрос.УстановитьПараметр("Ид", ИдПакета);
		В = Запрос.Выполнить().Выбрать();
		Если В.Следующий() Тогда
			Стр = Новый Структура("ВремяЗагрузки, Статус, СтрокДокументов, СуммаЦБ, СуммаБП",
				В.ВремяЗагрузки, В.Статус, В.СтрокДокументов, В.СуммаЦБ, В.СуммаБП);
			Возврат Стр;
		КонецЕсли;
	Исключение
	КонецПопытки;
	Возврат Неопределено;
КонецФункции

Функция _ДетПак_НайтиОшибки(ИдПакета)
	Результат = Новый Массив;
	Если Не ЗначениеЗаполнено(ИдПакета) Тогда Возврат Результат; КонецЕсли;
	Попытка
		Запрос = Новый Запрос(
			"ВЫБРАТЬ СообщениеОшибки, ТипОбъекта, КодОшибки, ЗначениеИсточника
			|ИЗ РегистрСведений.TL_ОшибкиЗагрузки
			|ГДЕ ИдентификаторПакета = &Ид
			|УПОРЯДОЧИТЬ ПО ВремяРегистрации УБЫВ");
		Запрос.УстановитьПараметр("Ид", ИдПакета);
		В = Запрос.Выполнить().Выбрать();
		Пока В.Следующий() Цикл
			Стр = Новый Структура("Сообщение, ТипОбъекта, КодОшибки, ЗначениеИсточника",
				В.СообщениеОшибки, Строка(В.ТипОбъекта), Строка(В.КодОшибки), В.ЗначениеИсточника);
			Результат.Добавить(Стр);
		КонецЦикла;
	Исключение
	КонецПопытки;
	Возврат Результат;
КонецФункции

Функция _ДетПак_ШаблонОшибка(Текст)
	Возврат "<!DOCTYPE html><html><body style=""font-family:'Segoe UI',sans-serif;padding:20px;"">"
		+ "<div style=""background:#FEE2E2;color:#991B1B;padding:16px;border-radius:6px;border-left:4px solid #EF4444;"">"
		+ Текст + "</div></body></html>";
КонецФункции

// === Служебные функции Markdown → HTML (H-направление) ===

// Проверить что символ — цифра.
Функция ИзЦифры(С)
	Возврат С >= "0" И С <= "9";
КонецФункции

// Проверить что строка начинается с "<цифра>. " (нумерованный список).
Функция НачинаетсяСНомера(С)
	ПозТочки = СтрНайти(С, ". ");
	Если ПозТочки = 0 ИЛИ ПозТочки > 4 Тогда
		Возврат Ложь;
	КонецЕсли;
	Часть = Лев(С, ПозТочки - 1);
	Для Сч = 1 По СтрДлина(Часть) Цикл
		Если НЕ ИзЦифры(Сред(Часть, Сч, 1)) Тогда
			Возврат Ложь;
		КонецЕсли;
	КонецЦикла;
	Возврат Истина;
КонецФункции

// Обработать инлайн-маркдаун в строке: **жирный**, *курсив*, `код`.
// Перед обработкой экранируем HTML-метасимволы.
Функция ОбработатьИнлайн(Знач Текст)
	Текст = Экр(Текст);
	// `код` → <code>код</code>
	Результат = "";
	ВнутриКода = Ложь;
	Буфер = "";
	Для Сч = 1 По СтрДлина(Текст) Цикл
		С = Сред(Текст, Сч, 1);
		Если С = "`" Тогда
			Если ВнутриКода Тогда
				Результат = Результат + "<code>" + Буфер + "</code>";
				Буфер = "";
				ВнутриКода = Ложь;
			Иначе
				Результат = Результат + ЗаменитьИнлайнФорматирование(Буфер);
				Буфер = "";
				ВнутриКода = Истина;
			КонецЕсли;
		Иначе
			Буфер = Буфер + С;
		КонецЕсли;
	КонецЦикла;
	Если ВнутриКода Тогда
		Результат = Результат + "`" + Буфер;
	Иначе
		Результат = Результат + ЗаменитьИнлайнФорматирование(Буфер);
	КонецЕсли;
	Возврат Результат;
КонецФункции

// Заменить **жирный** и *курсив* в строке (без обработки `код` — он уже обработан).
Функция ЗаменитьИнлайнФорматирование(Знач Текст)
	// Жирный **...**
	Текст = ЗаменитьПарный(Текст, "**", "<strong>", "</strong>");
	// Курсив *...*
	Текст = ЗаменитьПарный(Текст, "*", "<em>", "</em>");
	Возврат Текст;
КонецФункции

// Заменить парные маркеры в строке на открывающий/закрывающий теги.
Функция ЗаменитьПарный(Знач Текст, Маркер, ТегОткр, ТегЗакр)
	Результат = "";
	Поз = СтрНайти(Текст, Маркер);
	Открыт = Ложь;
	Пока Поз > 0 Цикл
		Результат = Результат + Лев(Текст, Поз - 1)
			+ ?(Открыт, ТегЗакр, ТегОткр);
		Текст = Сред(Текст, Поз + СтрДлина(Маркер));
		Открыт = НЕ Открыт;
		Поз = СтрНайти(Текст, Маркер);
	КонецЦикла;
	Результат = Результат + Текст;
	// Если остался открытый тег без закрывающего — добавляем маркер обратно
	// (упрощённое поведение — может оставить "сломанный" тег, но для простых статей ОК)
	Возврат Результат;
КонецФункции

// Закрыть открытые блоки (абзац / список / таблицу) при переходе к новому типу содержимого.
Функция ЗакрытьОткрытые(АбзацБуфер, ВСписке, ТипСписка, ВТаблице, ЗаголовокТаблицы, СтрокиТаблицы)
	Результат = "";
	Если ВТаблице Тогда
		Результат = Результат + ВывестиТаблицу(ЗаголовокТаблицы, СтрокиТаблицы);
	КонецЕсли;
	Если ВСписке Тогда
		Результат = Результат + "</" + ТипСписка + ">";
	КонецЕсли;
	Результат = Результат + ВыводитьАбзац(АбзацБуфер);
	Возврат Результат;
КонецФункции

// Вывести абзац (если не пустой).
Функция ВыводитьАбзац(Текст)
	Если ПустаяСтрока(Текст) Тогда
		Возврат "";
	КонецЕсли;
	Возврат "<p>" + ОбработатьИнлайн(Текст) + "</p>";
КонецФункции

// Вывести Markdown-таблицу. Шапка + строки тела (разделитель |--|--| опущен).
Функция ВывестиТаблицу(СтрокаШапки, СтрокиТела)
	HTML = "<table><thead><tr>";
	Для Каждого Я Из РазбитьЯчейкиТаблицы(СтрокаШапки) Цикл
		HTML = HTML + "<th>" + ОбработатьИнлайн(Я) + "</th>";
	КонецЦикла;
	HTML = HTML + "</tr></thead><tbody>";
	Для Каждого Р Из СтрокиТела Цикл
		HTML = HTML + "<tr>";
		Для Каждого Я Из РазбитьЯчейкиТаблицы(Р) Цикл
			HTML = HTML + "<td>" + ОбработатьИнлайн(Я) + "</td>";
		КонецЦикла;
		HTML = HTML + "</tr>";
	КонецЦикла;
	HTML = HTML + "</tbody></table>";
	Возврат HTML;
КонецФункции

// Разбить строку таблицы на ячейки. "| a | b |" → ["a", "b"].
Функция РазбитьЯчейкиТаблицы(Стр)
	Стр = СокрЛП(Стр);
	Если Лев(Стр, 1) = "|" Тогда Стр = Сред(Стр, 2); КонецЕсли;
	Если Прав(Стр, 1) = "|" Тогда Стр = Лев(Стр, СтрДлина(Стр) - 1); КонецЕсли;
	Ячейки = СтрРазделить(Стр, "|", Истина);
	Результат = Новый Массив;
	Для Каждого Я Из Ячейки Цикл
		Результат.Добавить(СокрЛП(Я));
	КонецЦикла;
	Возврат Результат;
КонецФункции

#КонецОбласти
