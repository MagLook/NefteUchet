"""
Генератор XML-метаданных для подсистемы TL_Сопутка (расширение TradeLedger v6.0).

Создаёт:
  Enums:
    TL_ТипОбъектаИсточника, TL_ПолеОшибки, TL_КодОшибкиЗагрузки, TL_СтатусПакета
  InformationRegisters:
    TL_СоответствиеИсточников, TL_ОшибкиЗагрузки, TL_СверкаПакета
  CommonModules (xml + копия .bsl из src):
    TL_МаппингЦБ, TL_Сверка, TL_СопуткаСервис

Все XML — UTF-8 BOM + CRLF, как требует 1С.
"""
import os
import shutil
import uuid

ROOT = r'D:\Users\magsp\ELSYPLUS\NefteUchet'
XML = os.path.join(ROOT, 'xml-v4')
SRC = os.path.join(ROOT, 'src', 'CommonModules')

XMLNS = (
    'xmlns="http://v8.1c.ru/8.3/MDClasses" '
    'xmlns:app="http://v8.1c.ru/8.2/managed-application/core" '
    'xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" '
    'xmlns:cmi="http://v8.1c.ru/8.2/managed-application/cmi" '
    'xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" '
    'xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" '
    'xmlns:style="http://v8.1c.ru/8.1/data/ui/style" '
    'xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" '
    'xmlns:v8="http://v8.1c.ru/8.1/data/core" '
    'xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" '
    'xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" '
    'xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" '
    'xmlns:xen="http://v8.1c.ru/8.3/xcf/enums" '
    'xmlns:xpr="http://v8.1c.ru/8.3/xcf/predef" '
    'xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" '
    'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
    'version="2.20"'
)

def u():
    return str(uuid.uuid4())

def write_xml(path, text):
    """UTF-8 BOM + CRLF."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    data = b'\xef\xbb\xbf' + text.replace('\n', '\r\n').encode('utf-8')
    with open(path, 'wb') as f:
        f.write(data)
    print(f'  wrote {os.path.relpath(path, ROOT)} ({len(data)} bytes)')

def write_bsl(path, src_bsl_name):
    """Копия из src — UTF-8 без BOM, LF (как у TL_ApiКлиент)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    src = os.path.join(SRC, src_bsl_name)
    shutil.copy(src, path)
    print(f'  copied {src_bsl_name} -> {os.path.relpath(path, ROOT)}')

# ============================================================================
#  COMMON MODULE
# ============================================================================

COMMON_MODULE_XML = '''<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject {ns}>
\t<CommonModule uuid="{uuid_main}">
\t\t<Properties>
\t\t\t<Name>{name}</Name>
\t\t\t<Synonym>
\t\t\t\t<v8:item>
\t\t\t\t\t<v8:lang>ru</v8:lang>
\t\t\t\t\t<v8:content>{synonym}</v8:content>
\t\t\t\t</v8:item>
\t\t\t</Synonym>
\t\t\t<Comment/>
\t\t\t<Global>false</Global>
\t\t\t<ClientManagedApplication>false</ClientManagedApplication>
\t\t\t<Server>true</Server>
\t\t\t<ExternalConnection>false</ExternalConnection>
\t\t\t<ClientOrdinaryApplication>false</ClientOrdinaryApplication>
\t\t\t<ServerCall>true</ServerCall>
\t\t\t<Privileged>false</Privileged>
\t\t\t<ReturnValuesReuse>DontUse</ReturnValuesReuse>
\t\t</Properties>
\t</CommonModule>
</MetaDataObject>'''

def gen_common_module(name, synonym):
    xml_path = os.path.join(XML, 'CommonModules', f'{name}.xml')
    bsl_path = os.path.join(XML, 'CommonModules', name, 'Ext', 'Module.bsl')
    text = COMMON_MODULE_XML.format(ns=XMLNS, uuid_main=u(), name=name, synonym=synonym)
    write_xml(xml_path, text)
    write_bsl(bsl_path, f'{name}.bsl')

# ============================================================================
#  ENUM
# ============================================================================

ENUM_XML = '''<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject {ns}>
\t<Enum uuid="{uuid_main}">
\t\t<InternalInfo>
\t\t\t<xr:GeneratedType name="EnumRef.{name}" category="Ref">
\t\t\t\t<xr:TypeId>{tid_ref}</xr:TypeId>
\t\t\t\t<xr:ValueId>{vid_ref}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t\t<xr:GeneratedType name="EnumManager.{name}" category="Manager">
\t\t\t\t<xr:TypeId>{tid_mgr}</xr:TypeId>
\t\t\t\t<xr:ValueId>{vid_mgr}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t\t<xr:GeneratedType name="EnumList.{name}" category="List">
\t\t\t\t<xr:TypeId>{tid_lst}</xr:TypeId>
\t\t\t\t<xr:ValueId>{vid_lst}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t</InternalInfo>
\t\t<Properties>
\t\t\t<Name>{name}</Name>
\t\t\t<Synonym>
\t\t\t\t<v8:item>
\t\t\t\t\t<v8:lang>ru</v8:lang>
\t\t\t\t\t<v8:content>{synonym}</v8:content>
\t\t\t\t</v8:item>
\t\t\t</Synonym>
\t\t\t<Comment/>
\t\t\t<UseStandardCommands>false</UseStandardCommands>
\t\t\t<Characteristics/>
\t\t\t<QuickChoice>true</QuickChoice>
\t\t\t<ChoiceMode>BothWays</ChoiceMode>
\t\t\t<DefaultListForm/>
\t\t\t<DefaultChoiceForm/>
\t\t\t<AuxiliaryListForm/>
\t\t\t<AuxiliaryChoiceForm/>
\t\t\t<ListPresentation/>
\t\t\t<ExtendedListPresentation/>
\t\t\t<Explanation/>
\t\t\t<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
\t\t</Properties>
\t\t<ChildObjects>
{values}
\t\t</ChildObjects>
\t</Enum>
</MetaDataObject>'''

ENUM_VALUE_XML = '''\t\t\t<EnumValue uuid="{uuid_val}">
\t\t\t\t<Properties>
\t\t\t\t\t<Name>{name}</Name>
\t\t\t\t\t<Synonym>
\t\t\t\t\t\t<v8:item>
\t\t\t\t\t\t\t<v8:lang>ru</v8:lang>
\t\t\t\t\t\t\t<v8:content>{synonym}</v8:content>
\t\t\t\t\t\t</v8:item>
\t\t\t\t\t</Synonym>
\t\t\t\t\t<Comment/>
\t\t\t\t</Properties>
\t\t\t</EnumValue>'''

def gen_enum(name, synonym, values):
    """values = [(name, synonym), ...]"""
    val_xml = '\n'.join(
        ENUM_VALUE_XML.format(uuid_val=u(), name=v_name, synonym=v_syn)
        for (v_name, v_syn) in values
    )
    text = ENUM_XML.format(
        ns=XMLNS,
        uuid_main=u(),
        tid_ref=u(), vid_ref=u(),
        tid_mgr=u(), vid_mgr=u(),
        tid_lst=u(), vid_lst=u(),
        name=name, synonym=synonym,
        values=val_xml,
    )
    path = os.path.join(XML, 'Enums', f'{name}.xml')
    write_xml(path, text)

# ============================================================================
#  INFORMATION REGISTER
# ============================================================================

REGISTER_XML = '''<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject {ns}>
\t<InformationRegister uuid="{uuid_main}">
\t\t<InternalInfo>
\t\t\t<xr:GeneratedType name="InformationRegisterRecord.{name}" category="Record">
\t\t\t\t<xr:TypeId>{u1a}</xr:TypeId><xr:ValueId>{u1b}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t\t<xr:GeneratedType name="InformationRegisterManager.{name}" category="Manager">
\t\t\t\t<xr:TypeId>{u2a}</xr:TypeId><xr:ValueId>{u2b}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t\t<xr:GeneratedType name="InformationRegisterSelection.{name}" category="Selection">
\t\t\t\t<xr:TypeId>{u3a}</xr:TypeId><xr:ValueId>{u3b}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t\t<xr:GeneratedType name="InformationRegisterList.{name}" category="List">
\t\t\t\t<xr:TypeId>{u4a}</xr:TypeId><xr:ValueId>{u4b}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t\t<xr:GeneratedType name="InformationRegisterRecordSet.{name}" category="RecordSet">
\t\t\t\t<xr:TypeId>{u5a}</xr:TypeId><xr:ValueId>{u5b}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t\t<xr:GeneratedType name="InformationRegisterRecordKey.{name}" category="RecordKey">
\t\t\t\t<xr:TypeId>{u6a}</xr:TypeId><xr:ValueId>{u6b}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t\t<xr:GeneratedType name="InformationRegisterRecordManager.{name}" category="RecordManager">
\t\t\t\t<xr:TypeId>{u7a}</xr:TypeId><xr:ValueId>{u7b}</xr:ValueId>
\t\t\t</xr:GeneratedType>
\t\t</InternalInfo>
\t\t<Properties>
\t\t\t<Name>{name}</Name>
\t\t\t<Synonym>
\t\t\t\t<v8:item>
\t\t\t\t\t<v8:lang>ru</v8:lang>
\t\t\t\t\t<v8:content>{synonym}</v8:content>
\t\t\t\t</v8:item>
\t\t\t</Synonym>
\t\t\t<Comment/>
\t\t\t<UseStandardCommands>true</UseStandardCommands>
\t\t\t<EditType>InDialog</EditType>
\t\t\t<DefaultRecordForm/>
\t\t\t<DefaultListForm/>
\t\t\t<AuxiliaryRecordForm/>
\t\t\t<AuxiliaryListForm/>
\t\t\t<InformationRegisterPeriodicity>Nonperiodical</InformationRegisterPeriodicity>
\t\t\t<WriteMode>Independent</WriteMode>
\t\t\t<MainFilterOnPeriod>false</MainFilterOnPeriod>
\t\t\t<IncludeHelpInContents>false</IncludeHelpInContents>
\t\t\t<DataLockControlMode>Managed</DataLockControlMode>
\t\t\t<FullTextSearch>DontUse</FullTextSearch>
\t\t\t<EnableTotalsSliceFirst>false</EnableTotalsSliceFirst>
\t\t\t<EnableTotalsSliceLast>false</EnableTotalsSliceLast>
\t\t\t<RecordPresentation/>
\t\t\t<ExtendedRecordPresentation/>
\t\t\t<ListPresentation/>
\t\t\t<ExtendedListPresentation/>
\t\t\t<Explanation/>
\t\t\t<DataHistory>DontUse</DataHistory>
\t\t\t<UpdateDataHistoryImmediatelyAfterWrite>false</UpdateDataHistoryImmediatelyAfterWrite>
\t\t\t<ExecuteAfterWriteDataHistoryVersionProcessing>false</ExecuteAfterWriteDataHistoryVersionProcessing>
\t\t</Properties>
\t\t<ChildObjects>
{children}
\t\t</ChildObjects>
\t</InformationRegister>
</MetaDataObject>'''

DIM_XML_HEAD = '''\t\t\t<Dimension uuid="{u}">
\t\t\t\t<Properties>
\t\t\t\t\t<Name>{name}</Name>
\t\t\t\t\t<Synonym>
\t\t\t\t\t\t<v8:item>
\t\t\t\t\t\t\t<v8:lang>ru</v8:lang>
\t\t\t\t\t\t\t<v8:content>{synonym}</v8:content>
\t\t\t\t\t\t</v8:item>
\t\t\t\t\t</Synonym>
\t\t\t\t\t<Comment/>
\t\t\t\t\t<Type>
{type_xml}
\t\t\t\t\t</Type>
\t\t\t\t\t<PasswordMode>false</PasswordMode>
\t\t\t\t\t<Format/>
\t\t\t\t\t<EditFormat/>
\t\t\t\t\t<ToolTip/>
\t\t\t\t\t<MarkNegatives>false</MarkNegatives>
\t\t\t\t\t<Mask/>
\t\t\t\t\t<MultiLine>false</MultiLine>
\t\t\t\t\t<ExtendedEdit>false</ExtendedEdit>
\t\t\t\t\t<MinValue xsi:nil="true"/>
\t\t\t\t\t<MaxValue xsi:nil="true"/>
\t\t\t\t\t<FillFromFillingValue>false</FillFromFillingValue>
\t\t\t\t\t<FillValue xsi:type="xs:string"/>
\t\t\t\t\t<FillChecking>DontCheck</FillChecking>
\t\t\t\t\t<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
\t\t\t\t\t<ChoiceParameterLinks/>
\t\t\t\t\t<ChoiceParameters/>
\t\t\t\t\t<QuickChoice>Auto</QuickChoice>
\t\t\t\t\t<CreateOnInput>Auto</CreateOnInput>
\t\t\t\t\t<ChoiceForm/>
\t\t\t\t\t<LinkByType/>
\t\t\t\t\t<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
\t\t\t\t\t<Master>false</Master>
\t\t\t\t\t<MainFilter>{mainfilter}</MainFilter>
\t\t\t\t\t<DenyIncompleteValues>false</DenyIncompleteValues>
\t\t\t\t\t<Indexing>DontIndex</Indexing>
\t\t\t\t\t<FullTextSearch>Use</FullTextSearch>
\t\t\t\t\t<DataHistory>Use</DataHistory>
\t\t\t\t\t<TypeReductionMode>TransformValues</TypeReductionMode>
\t\t\t\t</Properties>
\t\t\t</Dimension>'''

RES_XML = '''\t\t\t<Resource uuid="{u}">
\t\t\t\t<Properties>
\t\t\t\t\t<Name>{name}</Name>
\t\t\t\t\t<Synonym>
\t\t\t\t\t\t<v8:item>
\t\t\t\t\t\t\t<v8:lang>ru</v8:lang>
\t\t\t\t\t\t\t<v8:content>{synonym}</v8:content>
\t\t\t\t\t\t</v8:item>
\t\t\t\t\t</Synonym>
\t\t\t\t\t<Comment/>
\t\t\t\t\t<Type>
{type_xml}
\t\t\t\t\t</Type>
\t\t\t\t\t<PasswordMode>false</PasswordMode>
\t\t\t\t\t<Format/>
\t\t\t\t\t<EditFormat/>
\t\t\t\t\t<ToolTip/>
\t\t\t\t\t<MarkNegatives>false</MarkNegatives>
\t\t\t\t\t<Mask/>
\t\t\t\t\t<MultiLine>false</MultiLine>
\t\t\t\t\t<ExtendedEdit>false</ExtendedEdit>
\t\t\t\t\t<MinValue xsi:nil="true"/>
\t\t\t\t\t<MaxValue xsi:nil="true"/>
\t\t\t\t\t<FillFromFillingValue>false</FillFromFillingValue>
\t\t\t\t\t<FillValue xsi:type="xs:string"/>
\t\t\t\t\t<FillChecking>DontCheck</FillChecking>
\t\t\t\t\t<ChoiceFoldersAndItems>Items</ChoiceFoldersAndItems>
\t\t\t\t\t<ChoiceParameterLinks/>
\t\t\t\t\t<ChoiceParameters/>
\t\t\t\t\t<QuickChoice>Auto</QuickChoice>
\t\t\t\t\t<CreateOnInput>Auto</CreateOnInput>
\t\t\t\t\t<ChoiceForm/>
\t\t\t\t\t<LinkByType/>
\t\t\t\t\t<ChoiceHistoryOnInput>Auto</ChoiceHistoryOnInput>
\t\t\t\t\t<Indexing>DontIndex</Indexing>
\t\t\t\t\t<FullTextSearch>Use</FullTextSearch>
\t\t\t\t\t<DataHistory>Use</DataHistory>
\t\t\t\t</Properties>
\t\t\t</Resource>'''

def t_string(length):
    return f'\t\t\t\t\t\t<v8:Type>xs:string</v8:Type>\n\t\t\t\t\t\t<v8:StringQualifiers>\n\t\t\t\t\t\t\t<v8:Length>{length}</v8:Length>\n\t\t\t\t\t\t\t<v8:AllowedLength>Variable</v8:AllowedLength>\n\t\t\t\t\t\t</v8:StringQualifiers>'

def t_number(precision=15, scale=2):
    return f'\t\t\t\t\t\t<v8:Type>xs:decimal</v8:Type>\n\t\t\t\t\t\t<v8:NumberQualifiers>\n\t\t\t\t\t\t\t<v8:Digits>{precision}</v8:Digits>\n\t\t\t\t\t\t\t<v8:FractionDigits>{scale}</v8:FractionDigits>\n\t\t\t\t\t\t\t<v8:AllowedSign>Any</v8:AllowedSign>\n\t\t\t\t\t\t</v8:NumberQualifiers>'

def t_date(fractions='DateTime'):
    return f'\t\t\t\t\t\t<v8:Type>xs:dateTime</v8:Type>\n\t\t\t\t\t\t<v8:DateQualifiers>\n\t\t\t\t\t\t\t<v8:DateFractions>{fractions}</v8:DateFractions>\n\t\t\t\t\t\t</v8:DateQualifiers>'

def t_bool():
    return '\t\t\t\t\t\t<v8:Type>xs:boolean</v8:Type>'

def t_enum_ref(enum_name):
    return f'\t\t\t\t\t\t<v8:Type>cfg:EnumRef.{enum_name}</v8:Type>'

def t_anyref():
    return '\t\t\t\t\t\t<v8:TypeSet>cfg:AnyRef</v8:TypeSet>'

def make_dim(name, synonym, type_xml, mainfilter='true'):
    return DIM_XML_HEAD.format(u=u(), name=name, synonym=synonym, type_xml=type_xml, mainfilter=mainfilter)

def make_res(name, synonym, type_xml):
    return RES_XML.format(u=u(), name=name, synonym=synonym, type_xml=type_xml)

def gen_register(name, synonym, dims, resources):
    children = '\n'.join(resources + dims)
    text = REGISTER_XML.format(
        ns=XMLNS,
        uuid_main=u(),
        u1a=u(), u1b=u(), u2a=u(), u2b=u(), u3a=u(), u3b=u(),
        u4a=u(), u4b=u(), u5a=u(), u5b=u(), u6a=u(), u6b=u(), u7a=u(), u7b=u(),
        name=name, synonym=synonym,
        children=children,
    )
    path = os.path.join(XML, 'InformationRegisters', f'{name}.xml')
    write_xml(path, text)

# ============================================================================
#  ГЕНЕРАЦИЯ
# ============================================================================

print('=== Common Modules ===')
gen_common_module('TL_МаппингЦБ',     'Маппинг ЦБ')
gen_common_module('TL_Сверка',        'Сверка пакета')
gen_common_module('TL_СопуткаСервис', 'Сопутка сервис')

print('=== Enums ===')
gen_enum('TL_ТипОбъектаИсточника', 'Тип объекта источника', [
    ('Номенклатура',           'Номенклатура'),
    ('Поступление',            'Поступление'),
    ('ОтчётОРозничныхПродажах','Отчёт о розничных продажах'),
    ('ПКО',                    'ПКО'),
    ('Перемещение',            'Перемещение'),
    ('Списание',               'Списание'),
    ('Инвентаризация',         'Инвентаризация'),
    ('Оприходование',          'Оприходование'),
    ('Контрагент',             'Контрагент'),
    ('Склад',                  'Склад'),
    ('Организация',            'Организация'),
    ('Договор',                'Договор'),
    ('ВыпускБлюд',             'Выпуск блюд'),
    ('ТребованиеНакладная',    'Требование-накладная'),
])

gen_enum('TL_ПолеОшибки', 'Поле ошибки', [
    ('Номенклатура', 'Номенклатура'),
    ('Контрагент',   'Контрагент'),
    ('Склад',        'Склад'),
    ('Договор',      'Договор'),
    ('Организация',  'Организация'),
    ('СтавкаНДС',    'Ставка НДС'),
    ('Счёт',         'Счёт'),
    ('Сумма',        'Сумма'),
    ('Прочее',       'Прочее'),
])

gen_enum('TL_КодОшибкиЗагрузки', 'Код ошибки загрузки', [
    ('НСИ_НеНайдена',            'НСИ не найдена'),
    ('СтавкаНДС_НеСопоставлена', 'Ставка НДС не сопоставлена'),
    ('СчётУчёта_НеОпределён',    'Счёт учёта не определён'),
    ('СуммаНеСходится',          'Сумма не сходится'),
    ('ОбязательноеПолеПусто',    'Обязательное поле пусто'),
    ('ВидОперацииНеСопоставлен', 'Вид операции не сопоставлен'),
    ('ОшибкаПроведения',         'Ошибка проведения'),
    ('Прочее',                   'Прочее'),
])

gen_enum('TL_СтатусПакета', 'Статус пакета', [
    ('Получен',         'Получен'),
    ('Валидирован',     'Валидирован'),
    ('ПринятоЦеликом',  'Принято целиком'),
    ('ПринятоЧастично', 'Принято частично'),
    ('Отклонено',       'Отклонено'),
])

print('=== Information Registers ===')

# TL_СоответствиеИсточников
gen_register(
    'TL_СоответствиеИсточников', 'Соответствие источников',
    dims=[
        make_dim('Тип',          'Тип',          t_enum_ref('TL_ТипОбъектаИсточника')),
        make_dim('ИсточникUUID', 'Источник UUID', t_string(36)),
    ],
    resources=[
        make_res('СсылкаОбъект',                'Ссылка объект',           t_anyref()),
        make_res('ДатаПоследнейЗагрузки',       'Дата последней загрузки', t_date()),
        make_res('ИдентификаторПоследнегоПакета','Идентификатор пакета',   t_string(36)),
        make_res('ХешПоследнегоПакета',         'Хеш',                     t_string(64)),
    ],
)

# TL_ОшибкиЗагрузки
gen_register(
    'TL_ОшибкиЗагрузки', 'Ошибки загрузки',
    dims=[
        make_dim('ИдентификаторПакета', 'Идентификатор пакета', t_string(36)),
        make_dim('НомерЗаписи',         'Номер записи',         t_number(10, 0)),
    ],
    resources=[
        make_res('ИсточникUUID',     'Источник UUID',     t_string(36)),
        make_res('ТипОбъекта',       'Тип объекта',       t_enum_ref('TL_ТипОбъектаИсточника')),
        make_res('Поле',             'Поле',              t_enum_ref('TL_ПолеОшибки')),
        make_res('КодОшибки',        'Код ошибки',        t_enum_ref('TL_КодОшибкиЗагрузки')),
        make_res('ЗначениеИсточника','Значение источника',t_string(200)),
        make_res('СообщениеОшибки',  'Сообщение',         t_string(500)),
        make_res('ВремяРегистрации', 'Время регистрации', t_date()),
        make_res('Документ',         'Документ',          t_anyref()),
    ],
)

# TL_СверкаПакета
gen_register(
    'TL_СверкаПакета', 'Сверка пакета',
    dims=[
        make_dim('ИдентификаторПакета', 'Идентификатор пакета', t_string(36)),
    ],
    resources=[
        make_res('Источник',        'Источник',        t_string(50)),
        make_res('ВремяЗагрузки',   'Время загрузки',  t_date()),
        make_res('ПериодС',         'Период с',        t_date('Date')),
        make_res('ПериодПо',        'Период по',       t_date('Date')),
        make_res('СтрокНСИ',        'Строк НСИ',       t_number(10, 0)),
        make_res('СтрокДокументов', 'Строк документов',t_number(10, 0)),
        make_res('СуммаЦБ',         'Сумма ЦБ',        t_number(15, 2)),
        make_res('НДС_ЦБ',          'НДС ЦБ',          t_number(15, 2)),
        make_res('СуммаБП',         'Сумма БП',        t_number(15, 2)),
        make_res('НДС_БП',          'НДС БП',          t_number(15, 2)),
        make_res('Расхождение',     'Расхождение',     t_number(15, 2)),
        make_res('Статус',          'Статус',          t_enum_ref('TL_СтатусПакета')),
        make_res('Комментарий',     'Комментарий',     t_string(500)),
    ],
)

print('Done.')
