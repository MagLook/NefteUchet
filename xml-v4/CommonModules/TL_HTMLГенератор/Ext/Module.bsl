////////////////////////////////////////////////////////////////////////////////
// Общий модуль TL_HTMLГенератор
// Расширение TradeLedger для 1С:БП 3.0
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
Функция СформироватьДашборд(ВсегоСмен, Загружено, Новых, Ошибок, СуммаЗаПериод, СтатусSTS = "", ПоследнийЛог = "") Экспорт

	СуммаТекст = ФорматСуммы(СуммаЗаПериод);

	HTML = "<!DOCTYPE html><html><head><meta charset=""utf-8""></head>"
		+ "<body style=""margin:0; padding:4px 8px; font-family:-apple-system,'Segoe UI',sans-serif;"
		+ " background:#F8FAFC; font-size:13px;"">"
		// --- KPI карточки (одна строка) ---
		+ "<div style=""display:flex; gap:8px; margin-bottom:6px;"">"
		+ МиниКарточка(XMLСтрока(ВсегоСмен), "Смен", "#3B82F6")
		+ МиниКарточка(XMLСтрока(Загружено), "В 1С", "#16A34A")
		+ МиниКарточка(XMLСтрока(Новых), "Новых", "#EAB308")
		+ МиниКарточка(XMLСтрока(Ошибок), "Ошибок", "#EF4444")
		+ МиниКарточка(СуммаТекст, "Выручка", "#6B7280")
		+ "</div>"
		// --- Статус STS + последний лог ---
		+ "<div style=""display:flex; gap:12px; font-size:12px; color:#6B7280;"">";

	Если ЗначениеЗаполнено(СтатусSTS) Тогда
		HTML = HTML + "<span>" + Экр(СтатусSTS) + "</span>";
	КонецЕсли;

	Если ЗначениеЗаполнено(ПоследнийЛог) Тогда
		HTML = HTML + "<span style=""flex:1; text-align:right; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;"">"
			+ Экр(ПоследнийЛог) + "</span>";
	КонецЕсли;

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

	// ===== ГЕНЕРАЦИЯ HTML =====

	HTML = СтильСмены();

	// --- ШАПКА ---
	HTML = HTML + "<div class=""header""><div>"
		+ "<div class=""header-station"">" + Экр(Заголовок) + "</div>"
		+ "</div><div>"
		+ "<div class=""header-total"">" + Формат(ИтогоВсеСумма, "ЧДЦ=2; ЧРД=,; ЧГ=' '") + " &#8381;</div>"
		+ "<div class=""header-sub"">" + Формат(ИтогоВсеЛитры, "ЧДЦ=1") + " л &middot; "
		+ XMLСтрока(КолТоплива) + " вида топлива</div>"
		+ "</div></div>";

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

	// Итоги розницы для отдельной строки
	РозницаПоТопливу = Новый Соответствие;
	Для Каждого КодТ Из КодыТоплива Цикл
		РозницаПоТопливу.Вставить(КодТ, Новый Структура("Литры, Сумма", 0, 0));
	КонецЦикла;
	РозницаИтого = 0;
	ПеремещенияHTML = "";

	Для Каждого КЗ Из АгрегатОплат Цикл
		ИмяОпл = КЗ.Ключ;
		ДанныеОпл = КЗ.Значение;
		ИмяНРег = НРег(ИмяОпл);

		// Определить маршрут и проводку
		Если СтрНайти(ИмяНРег, "наличн") > 0 Тогда
			Маршрут = "<span class=""route-tag route-retail"">Розница</span>";
			Проводка = "Дт 50.01 Кт 62.Р";
			ЭтоРозница = Истина;
		ИначеЕсли СтрНайти(ИмяНРег, "сбербанк") > 0 Тогда
			Маршрут = "<span class=""route-tag route-retail"">Розница</span>";
			Проводка = "Дт 57.03 Кт 62.Р";
			ЭтоРозница = Истина;
		ИначеЕсли СтрНайти(ИмяНРег, "мобил") > 0 ИЛИ СтрНайти(ИмяНРег, "яндекс") > 0 Тогда
			Маршрут = "<span class=""route-tag route-transfer"">ЯНДЕКС</span>";
			Проводка = "Дт 41.02(ЯНДЕКС) Кт 41.02(АЗС)";
			ЭтоРозница = Ложь;
		ИначеЕсли СтрНайти(ИмяНРег, "корп") > 0 ИЛИ СтрНайти(ИмяНРег, "агора") > 0
			ИЛИ ИмяНРег = "кр" Тогда
			Маршрут = "<span class=""route-tag route-transfer"">Карты</span>";
			Проводка = "Дт 41.02(Карты) Кт 41.02(АЗС)";
			ЭтоРозница = Ложь;
		Иначе
			Маршрут = Экр(ИмяОпл);
			Проводка = "";
			ЭтоРозница = Ложь;
		КонецЕсли;

		СтрокаHTML = "<tr><td>" + Экр(ИмяОпл) + "</td><td>" + Маршрут + "</td>"
			+ "<td class=""provodka"">" + Проводка + "</td>";

		ИтогоСтроки = 0;
		Для Каждого КодТ Из КодыТоплива Цикл
			Яч = ДанныеОпл.Получить(КодТ);
			Если Яч <> Неопределено И Яч.Литры > 0 Тогда
				СтрокаHTML = СтрокаHTML + "<td class=""num"">" + Формат(Яч.Литры, "ЧДЦ=1")
					+ "</td><td class=""num"">" + Формат(Яч.Сумма, "ЧДЦ=0") + "</td>";
				ИтогоСтроки = ИтогоСтроки + Яч.Сумма;
				Если ЭтоРозница Тогда
					Р = РозницаПоТопливу.Получить(КодТ);
					Р.Литры = Р.Литры + Яч.Литры;
					Р.Сумма = Р.Сумма + Яч.Сумма;
				КонецЕсли;
			Иначе
				СтрокаHTML = СтрокаHTML + "<td class=""num"">&mdash;</td><td class=""num"">&mdash;</td>";
			КонецЕсли;
		КонецЦикла;
		СтрокаHTML = СтрокаHTML + "<td class=""num""><b>" + Формат(ИтогоСтроки, "ЧДЦ=0") + "</b></td></tr>";

		Если ЭтоРозница Тогда
			HTML = HTML + СтрокаHTML;
			РозницаИтого = РозницаИтого + ИтогоСтроки;
		Иначе
			ПеремещенияHTML = ПеремещенияHTML + СтрокаHTML;
		КонецЕсли;
	КонецЦикла;

	// Строка итого розницы
	HTML = HTML + "<tr class=""total-row""><td colspan=""3"">ИТОГО розница</td>";
	Для Каждого КодТ Из КодыТоплива Цикл
		Р = РозницаПоТопливу.Получить(КодТ);
		HTML = HTML + "<td class=""num"">" + Формат(Р.Литры, "ЧДЦ=1") + "</td>"
			+ "<td class=""num"">" + Формат(Р.Сумма, "ЧДЦ=0") + "</td>";
	КонецЦикла;
	HTML = HTML + "<td class=""num""><b>" + Формат(РозницаИтого, "ЧДЦ=0") + "</b></td></tr>";

	// Перемещения
	Если ЗначениеЗаполнено(ПеремещенияHTML) Тогда
		HTML = HTML + ПеремещенияHTML;
	КонецЕсли;

	// Строка итого за смену
	HTML = HTML + "<tr class=""total-row""><td colspan=""3"">ИТОГО за смену</td>";
	Для Каждого КодТ Из КодыТоплива Цикл
		Итог = ИтогиПоТопливу.Получить(КодТ);
		HTML = HTML + "<td class=""num"">" + Формат(Итог.Литры, "ЧДЦ=1") + "</td>"
			+ "<td class=""num"">" + Формат(Итог.Сумма, "ЧДЦ=0") + "</td>";
	КонецЦикла;
	HTML = HTML + "<td class=""num""><b>" + Формат(ИтогоВсеСумма, "ЧДЦ=0") + "</b></td></tr></table>";

	// --- ПРОВОДКИ ---
	HTML = HTML + "<table style=""margin-bottom:2px"">"
		+ "<tr><th colspan=""3"" style=""text-align:left"">Проводки при проведении</th></tr>"
		+ "<tr><td class=""provodka"" style=""width:220px"">Дт 62.Р Кт 90.01.1</td>"
		+ "<td>Выручка (по каждому виду топлива)</td>"
		+ "<td class=""num"">" + Формат(РозницаИтого, "ЧДЦ=0") + "</td></tr>"
		+ "<tr><td class=""provodka"">Дт 90.02.1 Кт 41.02</td>"
		+ "<td>Списание себестоимости с розничного склада</td>"
		+ "<td class=""num"">по учётной цене</td></tr>"
		+ "<tr><td class=""provodka"">Дт 50.01 Кт 62.Р</td>"
		+ "<td>Оплата наличными</td><td class=""num""></td></tr>"
		+ "<tr><td class=""provodka"">Дт 57.03 Кт 62.Р</td>"
		+ "<td>Оплата картой (эквайринг)</td><td class=""num""></td></tr>"
		+ "<tr><td class=""provodka"">Дт 90.03 Кт 68.02</td>"
		+ "<td>НДС 22%</td>"
		+ "<td class=""num"">" + Формат(Окр(РозницаИтого * 22 / 122, 0), "ЧДЦ=0") + "</td></tr>"
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
				+ "<span class=""tank-bar-label lbl-left"" style=""color:#fff"">" + Формат(ОстатокЛ, "ЧДЦ=0") + " л</span>"
				+ "<span class=""tank-bar-label lbl-right"">" + Формат(НачалоЛ, "ЧДЦ=0") + " л</span>"
				+ "</div>"
				+ "<table class=""tank-stats""><tr>"
				+ "<td>Отпущено</td><td class=""num negative"">&minus;" + Формат(РасходЛ, "ЧДЦ=0") + " л</td>"
				+ "<td>&rho;</td><td class=""num"">" + ?(ПлотнКон > 0, Формат(ПлотнКон, "ЧДЦ=3"), "-") + "</td>"
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
			HTML = HTML + "<div class=""" + Класс + """>"
				+ "<div class=""doc-check"">&#10003;</div>"
				+ "<div class=""doc-name""><b>" + Экр(Док.ТипДокумента) + "</b> " + Экр(Док.Описание) + "</div>"
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
		+ "<tr><td>Масса</td><td class=""num""><b>" + Формат(МассаТонн, "ЧДЦ=3") + " т</b> (" + Формат(МассаКг, "ЧДЦ=0") + " кг)</td></tr>"
		+ "<tr><td>Объём (книжный)</td><td class=""num""><b>" + Формат(Литры, "ЧДЦ=0") + " л</b></td></tr>"
		+ "<tr><td>Плотность</td><td class=""num"">" + ?(Плотность > 0, Формат(Плотность, "ЧДЦ=4") + " кг/л", "&mdash;") + "</td></tr>"
		+ "</table>";

	// --- ЦЕПОЧКА ДОКУМЕНТОВ ---
	HTML = HTML + "<div class=""section"">Документы для создания</div><div>";

	// Шаг 1: Перемещение
	HTML = HTML + "<div class=""doc-row checked"">"
		+ "<div class=""doc-check"">&#10003;</div>"
		+ "<div class=""doc-name""><b>Перемещение</b> Осн.склад &rarr; АЗС</div>"
		+ "<div class=""doc-detail"">Дт 41.01(АЗС) Кт 41.01(Основной) &middot; "
		+ Экр(ИмяТоплива) + " (т) " + Формат(МассаТонн, "ЧДЦ=3") + " т</div>"
		+ "<div class=""doc-amount"">" + Формат(МассаТонн, "ЧДЦ=3") + " т</div></div>";

	// Шаг 2: Комплектация
	HTML = HTML + "<div class=""doc-row checked"">"
		+ "<div class=""doc-check"">&#10003;</div>"
		+ "<div class=""doc-name""><b>Комплектация</b> тонны &rarr; литры</div>"
		+ "<div class=""doc-detail"">Дт 41.02 Кт 41.01 &middot; " + Экр(ИмяТоплива) + " (т) "
		+ Формат(МассаТонн, "ЧДЦ=3") + " т &rarr; " + Экр(ИмяТоплива) + " (л) "
		+ Формат(Литры, "ЧДЦ=0") + " л &middot; &rho;="
		+ ?(Плотность > 0, Формат(Плотность, "ЧДЦ=4"), "-") + "</div>"
		+ "<div class=""doc-amount"">" + Формат(Литры, "ЧДЦ=0") + " л</div></div>";

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
	|h3 { color:#3B82F6; font-size:13px; text-transform:uppercase; letter-spacing:1px; margin:16px 0 8px; }
	|.card { background:#fff; border-radius:12px; padding:16px; margin-bottom:12px;
	|  box-shadow:0 1px 3px rgba(0,0,0,0.08); }
	|.row { display:flex; gap:12px; margin-bottom:8px; align-items:center; }
	|.row label { min-width:140px; color:#6B7280; font-size:12px; }
	|.row input, .row select { flex:1; padding:6px 10px; border:1px solid #D9DFE7;
	|  border-radius:6px; font-size:13px; }
	|.row input:focus { outline:none; border-color:#3B82F6; box-shadow:0 0 0 2px rgba(59,130,246,0.15); }
	|table { width:100%; border-collapse:collapse; }
	|th { text-align:left; font-size:11px; color:#6B7280; text-transform:uppercase;
	|  padding:4px 8px; border-bottom:2px solid #D9DFE7; }
	|td { padding:4px 8px; border-bottom:1px solid #D9DFE7; }
	|td input { width:100%; padding:4px 6px; border:1px solid #D9DFE7; border-radius:4px; font-size:12px; }
	|.btn { padding:8px 24px; border:none; border-radius:8px; font-size:13px;
	|  font-weight:600; cursor:pointer; }
	|.btn-primary { background:#3B82F6; color:#fff; }
	|.btn-primary:hover { background:#2563EB; }
	|.btn-secondary { background:#E5E7EB; color:#374151; }
	|.status { margin-top:8px; padding:8px; border-radius:6px; display:none; }
	|.status.ok { display:block; background:#F0FDF4; color:#16A34A; }
	|.status.err { display:block; background:#FEF2F2; color:#EF4444; }
	|</style></head><body>
	|<h2>Настройки TradeLedger</h2>
	|
	|<div class=""card"">
	|<h3>Подключение к API</h3>
	|<div class=""row""><label>URL сервера</label><input id=""s_url"" value=""""></div>
	|<div class=""row""><label>Логин</label><input id=""s_login"" value=""""></div>
	|<div class=""row""><label>Пароль</label><input id=""s_password"" type=""password"" value=""""></div>
	|<div class=""row""><label>Код системы</label><input id=""s_system"" value="""" style=""max-width:100px;""></div>
	|</div>
	|
	|<div class=""card"">
	|<h3>Организация</h3>
	|<div class=""row""><label>Организация</label><input id=""s_org"" value=""""></div>
	|<div class=""row""><label>Основной склад</label><input id=""s_warehouse"" value=""""></div>
	|<div class=""row""><label>Поставщик</label><input id=""s_supplier"" value=""""></div>
	|</div>
	|
	|<div class=""card"">
	|<h3>Станции</h3>
	|<table>
	|<tr><th>Код</th><th>Наименование</th><th>Склад 1С</th><th>Закр.смены</th></tr>
	|<tbody id=""tblStations""></tbody>
	|</table>
	|<button class=""btn btn-secondary"" onclick=""addStation()"" style=""margin-top:8px;"">+ Станция</button>
	|</div>
	|
	|<div class=""card"">
	|<h3>Топливо</h3>
	|<table>
	|<tr><th>Код</th><th>Название</th><th>Номенклатура (т)</th><th>Номенклатура (л)</th><th>Плотность</th></tr>
	|<tbody id=""tblFuel""></tbody>
	|</table>
	|<button class=""btn btn-secondary"" onclick=""addFuel()"" style=""margin-top:8px;"">+ Топливо</button>
	|</div>
	|
	|<div class=""card"">
	|<h3>Каналы оплат</h3>
	|<table>
	|<tr><th>Ключ</th><th>Наименование</th><th>Склад</th><th>Перемещение</th></tr>
	|<tbody id=""tblPayments""></tbody>
	|</table>
	|</div>
	|
	|<div style=""display:flex; gap:12px; margin-top:16px;"">
	|<button class=""btn btn-primary"" id=""btnSave"">Сохранить</button>
	|<div id=""statusMsg"" class=""status""></div>
	|</div>
	|
	|<script>
	|var cfg = {};
	|try { cfg = JSON.parse('" + Экр(ТекущиеНастройки) + "'); } catch(e) { cfg = {}; }
	|
	|function v(key, def) { return cfg[key] || def || ''; }
	|
	|document.getElementById('s_url').value = v('URLСервера');
	|document.getElementById('s_login').value = v('Логин');
	|document.getElementById('s_password').value = v('Пароль');
	|document.getElementById('s_system').value = v('КодСистемы');
	|document.getElementById('s_org').value = v('Организация');
	|document.getElementById('s_warehouse').value = v('ОсновнойСклад');
	|document.getElementById('s_supplier').value = v('Поставщик');
	|
	|// Станции
	|var stations = [];
	|Object.keys(cfg).forEach(function(k) {
	|  var m = k.match(/^Станция_(\d+)_Наименование$/);
	|  if (m) stations.push(m[1]);
	|});
	|if (stations.length === 0) stations = ['5'];
	|stations.forEach(function(code) { addStation(code); });
	|
	|function addStation(code) {
	|  code = code || '';
	|  var tb = document.getElementById('tblStations');
	|  var tr = document.createElement('tr');
	|  tr.innerHTML = '<td><input class=""st_code"" value=""'+code+'"" style=""width:60px;""></td>'
	|    + '<td><input class=""st_name"" value=""'+v('Станция_'+code+'_Наименование')+'""></td>'
	|    + '<td><input class=""st_wh"" value=""'+v('Станция_'+code+'_Склад')+'""></td>'
	|    + '<td><input class=""st_close"" value=""'+v('Станция_'+code+'_ВремяЗакрытия','00:00')+'"" style=""width:70px;""></td>';
	|  tb.appendChild(tr);
	|}
	|
	|// Топливо
	|var fuels = [];
	|Object.keys(cfg).forEach(function(k) {
	|  var m = k.match(/^Топливо_(\w+)_Тонны$/);
	|  if (m) fuels.push(m[1]);
	|});
	|if (fuels.length === 0) fuels = ['2','3','5'];
	|var fuelNames = {'2':'АИ-92','3':'АИ-95','5':'ДТ','100':'АИ-100','98':'АИ-98'};
	|fuels.forEach(function(code) { addFuel(code); });
	|
	|function addFuel(code) {
	|  code = code || '';
	|  var tb = document.getElementById('tblFuel');
	|  var tr = document.createElement('tr');
	|  tr.innerHTML = '<td><input class=""f_code"" value=""'+code+'"" style=""width:60px;""></td>'
	|    + '<td><input class=""f_name"" value=""'+(fuelNames[code]||'')+'""></td>'
	|    + '<td><input class=""f_tons"" value=""'+v('Топливо_'+code+'_Тонны')+'""></td>'
	|    + '<td><input class=""f_litres"" value=""'+v('Топливо_'+code+'_Литры')+'""></td>'
	|    + '<td><input class=""f_density"" value=""'+v('Топливо_'+code+'_Плотность')+'"" style=""width:70px;""></td>';
	|  tb.appendChild(tr);
	|}
	|
	|// Оплаты
	|var payments = [
	|  {key:'retail', name:v('Оплата_retail_Наименование','Розница')},
	|  {key:'cards', name:v('Оплата_cards_Наименование','Карты')},
	|  {key:'online', name:v('Оплата_online_Наименование','Онлайн')},
	|  {key:'ledger', name:v('Оплата_ledger_Наименование','Ведомости')}
	|];
	|var ptb = document.getElementById('tblPayments');
	|payments.forEach(function(p) {
	|  var tr = document.createElement('tr');
	|  var req = v('Оплата_'+p.key+'_ТребуетПеремещения','Нет');
	|  tr.innerHTML = '<td>'+p.key+'</td>'
	|    + '<td><input class=""p_name"" data-key=""'+p.key+'"" value=""'+p.name+'""></td>'
	|    + '<td><input class=""p_wh"" data-key=""'+p.key+'"" value=""'+v('Оплата_'+p.key+'_Склад')+'""></td>'
	|    + '<td><select class=""p_move"" data-key=""'+p.key+'""><option'+(req==='Да'?' selected':'')+'>Да</option><option'+(req!=='Да'?' selected':'')+'>Нет</option></select></td>';
	|  ptb.appendChild(tr);
	|});
	|
	|// Сбор данных и сохранение
	|document.getElementById('btnSave').addEventListener('click', function() {
	|  var result = {};
	|  result['URLСервера'] = document.getElementById('s_url').value;
	|  result['Логин'] = document.getElementById('s_login').value;
	|  result['Пароль'] = document.getElementById('s_password').value;
	|  result['КодСистемы'] = document.getElementById('s_system').value;
	|  result['Организация'] = document.getElementById('s_org').value;
	|  result['ОсновнойСклад'] = document.getElementById('s_warehouse').value;
	|  result['Поставщик'] = document.getElementById('s_supplier').value;
	|
	|  // Станции
	|  var rows = document.getElementById('tblStations').rows;
	|  for (var i = 0; i < rows.length; i++) {
	|    var code = rows[i].querySelector('.st_code').value;
	|    if (!code) continue;
	|    result['Станция_'+code+'_Наименование'] = rows[i].querySelector('.st_name').value;
	|    result['Станция_'+code+'_Склад'] = rows[i].querySelector('.st_wh').value;
	|    result['Станция_'+code+'_ВремяЗакрытия'] = rows[i].querySelector('.st_close').value;
	|  }
	|
	|  // Топливо
	|  rows = document.getElementById('tblFuel').rows;
	|  for (var i = 0; i < rows.length; i++) {
	|    var code = rows[i].querySelector('.f_code').value;
	|    if (!code) continue;
	|    result['Топливо_'+code+'_Тонны'] = rows[i].querySelector('.f_tons').value;
	|    result['Топливо_'+code+'_Литры'] = rows[i].querySelector('.f_litres').value;
	|    result['Топливо_'+code+'_Плотность'] = rows[i].querySelector('.f_density').value;
	|  }
	|
	|  // Оплаты
	|  document.querySelectorAll('.p_name').forEach(function(el) {
	|    var key = el.dataset.key;
	|    result['Оплата_'+key+'_Наименование'] = el.value;
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
		+ "<div style=""font-size:36px;"">📋</div>"
		+ "<div style=""margin-top:8px;"">Выберите строку для просмотра деталей</div>"
		+ "</div>"
	);
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
	Возврат "<div style=""flex:1; background:#fff; border-left:3px solid " + Цвет + ";"
		+ " padding:6px 10px; border-radius:8px; box-shadow:0 1px 2px rgba(0,0,0,0.06);"">"
		+ "<span style=""font-size:18px; font-weight:700; color:" + Цвет + ";"">" + Значение + "</span>"
		+ " <span style=""font-size:11px; color:#9CA3AF;"">" + Подпись + "</span>"
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

#КонецОбласти
