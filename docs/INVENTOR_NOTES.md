# Паспорт возможностей: мост Python ↔ Autodesk Inventor 2027

Живой документ. Пополняется по мере тестирования COM API. Каждый раздел — что подтверждено рабочим способом, с конкретным кодом.

## Подключение

Рабочий и надёжный способ — **чистый динамический (late-bound) dispatch**:

```python
import win32com.client
app = win32com.client.GetActiveObject("Inventor.Application")
```

Требует уже запущенного Inventor на этой машине (COM работает только локально).

## ⚠️ Грабли: НЕ использовать gencache.EnsureDispatch / CastTo

Попытка получить "нормальные" именованные константы через раннее связывание:

```python
# ТАК ДЕЛАТЬ НЕЛЬЗЯ — ломает GetActiveObject на этой машине/версии Inventor 2027
app = win32com.client.gencache.EnsureDispatch(win32com.client.GetActiveObject("Inventor.Application"))
```

После первого же вызова `gencache.EnsureDispatch` пакет pywin32 генерирует и кэширует на диск
(`%TEMP%\gen_py\3.11\`) обёртку типовой библиотеки Inventor. Начиная с этого момента **даже
обычный `GetActiveObject` в новом процессе Python начинает падать**:

```
KeyError: '_dispobj_'
```

Также `win32com.client.CastTo(doc, "PartDocument")` требует ту же кэшированную обёртку и
приводит к той же проблеме.

**Фикс, если это случилось:** удалить кэш и вернуться к чистому dynamic dispatch:

```python
import shutil
shutil.rmtree(r"C:\Users\User\AppData\Local\Temp\gen_py", ignore_errors=True)
```

**Вывод:** во всём мосте используем только чистый dynamic dispatch. Именованные константы
(`kPartDocumentObject` и т.д.) не читаем из `win32com.client.constants` — используем захардкоженные
числа (см. таблицу ниже). При позднем связывании доступ к любым свойствам объекта (например,
`doc.ComponentDefinition`) работает без кастинга — pywin32 сам резолвит имена через `IDispatch`.

## Константы Inventor API (проверено на этой версии — 2027)

| Имя | Значение | Где используется |
|---|---|---|
| `kPartDocumentObject` | 12290 | `Documents.Add(...)` для детали |
| `kAssemblyDocumentObject` | 12291 | `Documents.Add(...)` для сборки |
| `kDrawingDocumentObject` | 12292 | `Documents.Add(...)` для чертежа |
| `kMillimeterLengthUnits` | 11269 | единицы длины |
| `kNewBodyOperation` | 20485 | экструзия: новое тело |
| `kJoinOperation` | 20481 | экструзия: объединение |
| `kCutOperation` | 20482 | экструзия: вычитание |
| `kPositiveExtentDirection` | 20993 | направление экструзии, "вперёд" |
| `kNegativeExtentDirection` | 20994 | направление экструзии, "назад" |
| `kSymmetricExtentDirection` | 20995 | направление экструзии, симметрично |

Остальные константы (сборка/крепления, виды чертежа, листовой металл и т.д.) добавляются сюда по
мере подтверждения — не гадать заранее.

### Как узнать новую константу — правильный способ (`tools/enum_lookup.py`)

Первая идея (`gencache.EnsureModule`/`EnsureDispatch` + чтение `win32com.client.constants`) —
**плохая**: сама генерация модуля на диске (даже без вызова `GetActiveObject` через неё) рано или
поздно ломает последующие `GetActiveObject` ошибкой `KeyError: '_dispobj_'`. Каждый раз пришлось бы
чистить `%TEMP%\gen_py`. К тому же `win32com.client.constants` не всегда наполняется даже после
`EnsureModule` (нужен ещё и `EnsureDispatch` на живом объекте, что усугубляет поломку).

**Рабочий способ — чистое чтение typelib через `pythoncom`, без gencache вообще:**

```python
ti = app._oleobj_.GetTypeInfo()
tlib, idx = ti.GetContainingTypeLib()
# перебор всех типов библиотеки, enum — это TKIND_ENUM = 0 (НЕ 3!)
for i in range(tlib.GetTypeInfoCount()):
    if tlib.GetTypeInfoType(i) != 0:
        continue
    name = tlib.GetDocumentation(i)[0]
    if name == "ViewOrientationTypeEnum":
        enum_ti = tlib.GetTypeInfo(i)
        cVars = enum_ti.GetTypeAttr()[7]
        for v in range(cVars):
            vd = enum_ti.GetVarDesc(v)
            varname = enum_ti.GetNames(vd[0])[0]
            print(varname, "=", vd[1])
```

Готовый инструмент: `bridge/tools/enum_lookup.py`. Использование:
```bash
python tools/enum_lookup.py ViewOrientationTypeEnum DrawingViewStyleEnum
```
Полностью безопасно мешать с обычным `GetActiveObject` в той же или соседних сессиях — не трогает
кэш вообще.

Тем же способом (`tools/reflect.py`) читаются **сигнатуры методов** любой COM-коллекции
(имена параметров, обязательные/опциональные, и — если параметр является ссылкой на другой
интерфейс — имя этого интерфейса, например `EdgeCollection`). Так был найден рабочий рецепт для
`FlangeFeatures.CreateFlangeDefinition` (см. раздел «Листовой металл» ниже) — метод ожидал
конкретно `EdgeCollection`, а не общий `ObjectCollection`.

## Пути шаблонов (эта установка Inventor 2027, локаль ru-RU)

```
C:\Users\Public\Documents\Autodesk\Inventor 2027\Templates\ru-RU\Metric\Standard (mm).ipt
C:\Users\Public\Documents\Autodesk\Inventor 2027\Templates\ru-RU\Metric\Sheet Metal (mm).ipt
C:\Users\Public\Documents\Autodesk\Inventor 2027\Templates\ru-RU\Metric\Standard (mm).iam
C:\Users\Public\Documents\Autodesk\Inventor 2027\Templates\ru-RU\Metric\ISO.idw
```
Папку шаблонов активного проекта можно получить программно:
```python
app.DesignProjectManager.ActiveDesignProject.TemplatesPath
```

## Документы

- Создание: `app.Documents.Add(kPartDocumentObject, template_path, True)`
- Открытые несохранённые документы остаются в `app.Documents` между запусками скриптов —
  **закрывать за собой** (`doc.Close(True)`), иначе накапливается мусор в сессии Inventor.
- Список открытых: `for d in app.Documents: d.DisplayName, d.FullFileName`

## Параметры

- Чтение: `doc.ComponentDefinition.Parameters.Item(name).Value / .Expression / .Units`
- Запись: `p.Expression = "25 mm"; doc.Update()`
- Работает через чистый dynamic dispatch без кастинга.

## iLogic

**Подтверждено рабочим тестом end-to-end** — правило создано программно, выполнено программно,
и **увидено и подтверждено в реальном окне «Пульта»** (webview-приложение): вкладка
«iLogic-правила» показала `SetThickness20`, толщина в модели реально изменилась 10мм → 20мм.

```python
addin = app.ApplicationAddIns.ItemById("{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}")
addin.Activate()
auto = addin.Automation

auto.AddRule(doc, "SetThickness20", 'Parameter("Thickness") = "20 mm"')  # создать правило
auto.RunRule(doc, "SetThickness20")                                      # выполнить
doc.Update()
```

- Список правил: `auto.Rules(doc)` → коллекция объектов с `.Name` (используется в `core.py`).
- `_oleobj_.GetTypeInfo()` на этом объекте **не работает** (`OLE error 0x80131165`) — значит
  `tools/reflect.py` тут бесполезен, имена методов пришлось брать из общеизвестной документации
  iLogic Automation API, не из живой интроспекции. Сработали с первого раза: `AddRule`, `RunRule`.

## Эскизы / экструзия

**Подтверждено рабочим тестом** — создан блок 100×60×10 мм, визуально проверен скриншотом
реального окна Inventor (виден в дереве модели: `Выдавливание1`, `Твердые тела(1)`).

```python
tg = app.TransientGeometry
xyPlane = compDef.WorkPlanes.Item(3)          # для Standard (mm).ipt — плоскость XY
sketch = compDef.Sketches.Add(xyPlane)
p1 = tg.CreatePoint2d(0, 0)
p2 = tg.CreatePoint2d(length_cm, width_cm)     # ВНИМАНИЕ: единицы — см, не мм (см. ниже)
rect = sketch.SketchLines.AddAsTwoPointRectangle(p1, p2)

profile = sketch.Profiles.AddForSolid()
extrudeDef = compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(profile, kNewBodyOperation)
extrudeDef.SetDistanceExtent(thickness_cm, kPositiveExtentDirection)
extrude = compDef.Features.ExtrudeFeatures.Add(extrudeDef)
```

### ⚠️ Единицы измерения: внутренний API Inventor всегда в САНТИМЕТРАХ

`Parameter.Value` для параметра с выражением `"100 mm"` вернёт `10.0` (см), а не `100`.
`TransientGeometry.CreatePoint2d`, `SetDistanceExtent` и т.д. тоже ждут сантиметры.
Проверено эмпирически — не полагаться на интуицию про мм.

### Не получилось (пока): жёсткая привязка distance экструзии к параметру

`extrude.ExtentOne.Distance.Expression = "Thickness"` → `AttributeError`. Похоже, свойство
`ExtentOne` не резолвится через чистый dynamic dispatch на объекте, возвращаемом
`ExtrudeFeatures.Add(...)`. Деталь при этом создаётся корректно с нужным размером (просто
экструзия — с фиксированным числом, не live-ссылкой на User Parameter). Нужно исследовать
отдельно, если понадобится живая перепривязка размера экструзии к параметру после создания.

### Полезно: скриншот реального окна Inventor для визуальной проверки

COM не даёт "photo", но можно снять весь экран через PowerShell/System.Drawing и подтвердить
результат глазами:

```powershell
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
# ShowWindow + SetForegroundWindow по MainWindowHandle процесса Inventor.exe, затем
# Graphics.CopyFromScreen в Bitmap, .Save(png)
```

## Листовой металл (гибка)

**Подтверждено рабочим тестом** — создана база (плоская грань) + один загиб (флэнч 90°),
проверено скриншотом реального окна Inventor (дерево модели: `Грань1`, `Фланец1`).

```python
tpl = r"...\Sheet Metal (mm).ipt"
doc = app.Documents.Add(12290, tpl, True)          # тот же kPartDocumentObject
compDef = doc.ComponentDefinition                    # автоматически SheetMetalComponentDefinition
compDef.Thickness.Value                              # толщина металла берётся из шаблона (см)

# базовая грань — из эскиза, через объект-definition (как и everywhere в Part API):
sketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3))
sketch.SketchLines.AddAsTwoPointRectangle(tg.CreatePoint2d(0,0), tg.CreatePoint2d(8,5))
profile = sketch.Profiles.AddForSolid()
faceDef = compDef.Features.FaceFeatures.CreateFaceFeatureDefinition(profile)
faceFeature = compDef.Features.FaceFeatures.Add(faceDef)

# загиб (флэнч) от ребра готовой грани:
edge = faceFeature.Faces.Item(1).Edges.Item(1)
edgeColl = app.TransientObjects.CreateEdgeCollection()   # НЕ CreateObjectCollection!
edgeColl.Add(edge)
import math
flangeDef = compDef.Features.FlangeFeatures.CreateFlangeDefinition(
    edgeColl, math.radians(90), 3)                        # угол в радианах, дистанция в см
flangeFeature = compDef.Features.FlangeFeatures.Add(flangeDef)
```

### ⚠️ Грабли: `Edges` параметр требует именно `EdgeCollection`, не `ObjectCollection`

`app.TransientObjects.CreateObjectCollection()` даёт объект, несовместимый по типу — падает с
`Несовместимые данные`. Нужен специальный `app.TransientObjects.CreateEdgeCollection()`
(аналогично: `CreateFaceCollection()` для граней). Найдено через `tools/reflect.py` —
ELEMDESC параметра указал точный тип интерфейса (`EdgeCollection`).

## Чертежи

**Подтверждено рабочим тестом** — создан лист ISO с рамкой/штампом и изометрическим видом
детали, проверено скриншотом (полноценный чертёж, вид с тенями, рамка с датой/автором).

```python
part_doc = app.Documents.Open(part_path, True)   # деталь должна быть открыта
drw = app.Documents.Add(kDrawingDocumentObject, r"...\ISO.idw", True)
sheet = drw.ActiveSheet
pos = tg.CreatePoint2d(10, 15)   # позиция вида на листе, см
view = sheet.DrawingViews.AddBaseView(
    part_doc, pos, 1,                                  # 1 = масштаб 1:1
    kIsoTopRightViewOrientation,                        # 10759
    kShadedDrawingViewStyle)                             # 32259
drw.SaveAs(r"...\TestDrawing.idw", False)
```

Шаблоны: `ISO.idw` (нейтральный), `GOST.idw` (для РФ) — оба лежат в той же папке `Metric\`.
Полный список констант `ViewOrientationTypeEnum` / `DrawingViewStyleEnum` — в
`tools/enum_lookup.py ViewOrientationTypeEnum DrawingViewStyleEnum`.

## Сборки

**Подтверждено рабочим тестом** — собраны два экземпляра `TestPart.ipt` в новой сборке, разнесены
по X, проверено скриншотом (дерево: `TestPart:1`, `TestPart:2`).

```python
doc = app.Documents.Add(kAssemblyDocumentObject, r"...\Standard (mm).iam", True)
compDef = doc.ComponentDefinition

m1 = tg.CreateMatrix()                                    # единичная матрица — в начале координат
occ1 = compDef.Occurrences.Add(part_path, m1)

m2 = tg.CreateMatrix()
m2.SetTranslation(tg.CreateVector(20, 0, 0))               # смещение 20 см по X
occ2 = compDef.Occurrences.Add(part_path, m2)
```

### Сопряжения (constraints) — сигнатуры найдены, живой тест не проводился

Через `tools/reflect.py` на `compDef.Constraints` найдены точные методы (не пришлось гадать):

```
AddMateConstraint(EntityOne, EntityTwo, Offset, EntityOneInferredType, EntityTwoInferredType, ...)
AddFlushConstraint(EntityOne, EntityTwo, Offset, ...)
AddInsertConstraint(EntityOne, EntityTwo, AxesOpposed, Distance, ...)
AddAngleConstraint(EntityOne, EntityTwo, Angle, SolutionType, ...)
AddTangentConstraint(...)
```

`EntityOne`/`EntityTwo` — конкретные грани/рёбра/оси **в контексте сборки** (через
`occurrence.CreateGeometryProxy(face)`, не сырые Face объекты детали). Не протестировано вживую —
следующий шаг, если понадобятся автоматические сопряжения.

## Экспорт / сохранение

**Подтверждено** — прямой `SaveAs` реально конвертирует форматы, отдельный Translator API не
понадобился:

- `doc.Save()`
- `doc.SaveAs(target_path, True)` — второй аргумент `True` = "save copy as" (не переключает
  активный документ на новый путь). Формат экспорта определяется по расширению целевого пути.
- Проверено: `.pdf` (чертёж → PDF, 118КБ, открывается), `.step` (деталь → STEP, 12.7КБ) — оба
  создались корректно с первого раза.

## iProperties, материал, массовые характеристики

**Подтверждено рабочим тестом** — включая прямую связку с обозначением из `NAMING.md`.

```python
dt = doc.PropertySets.Item("Design Tracking Properties")
dt.Item("Part Number").Value = "1000.0002.0004"      # то самое обозначение PDM
dt.Item("Description").Value = "Тестовая пластина"

summary = doc.PropertySets.Item("Summary Information")
summary.Item("Title").Value = "..."
summary.Item("Author").Value = "..."

# материал — назначается объектом из коллекции документа, не строкой:
for m in doc.Materials:
    if "Сталь" in m.Name:
        doc.ComponentDefinition.Material = m
        break

mp = doc.ComponentDefinition.MassProperties
mp.Mass            # кг
mp.Volume          # см³
mp.Area            # см²
mp.CenterOfMass.X/Y/Z
```

Все значения численно сошлись с геометрией детали (объём 10×6×1см = 60см³ — точно совпало).

## Дополнительные фичи детали: скругление и отверстие

**Подтверждено рабочим тестом**, включая одну реальную (не выдуманную) проблему и её починку.

```python
# скругление — тот же паттерн EdgeCollection, что и у флэнча:
edgeColl = app.TransientObjects.CreateEdgeCollection()
edgeColl.Add(body.Edges.Item(1))
compDef.Features.FilletFeatures.AddSimple(
    edgeColl, radius_cm, False, False, True, False, False, False)

# отверстие — точка эскиза на грани, через ObjectCollection (НЕ EdgeCollection):
sketch = compDef.Sketches.Add(topFace)
pt = sketch.SketchPoints.Add(tg.CreatePoint2d(x, y))
ptColl = app.TransientObjects.CreateObjectCollection()
ptColl.Add(pt)
placementDef = compDef.Features.HoleFeatures.CreateSketchPlacementDefinition(ptColl)
compDef.Features.HoleFeatures.AddDrilledByThroughAllExtent(
    placementDef, diameter_cm, kPositiveExtentDirection)   # 20993 — см. ниже про направление!
```

### ⚠️ Грабли: неправильное направление отверстия — тихий отказ без ошибки

С `kNegativeExtentDirection` (20994) COM-вызов отработал **без единой ошибки**, `HealthStatus`
феатуры был "здоровым" на вид, но объём тела не изменился — отверстие физически не резало
материал (бурило "в пустоту", в сторону от тела). Обнаружено только сравнением
`MassProperties.Volume` до/после и числа граней тела (`SurfaceBodies.Item(1).Faces.Count`
не увеличилось). При правильном `kPositiveExtentDirection` (20993) грань добавилась (+1) и объём
уменьшился ровно на ожидаемые `π·r²·h`.

**Вывод: после создания feature, которая должна менять геометрию, всегда проверять реальный
эффект (объём/число граней), а не только отсутствие исключения** — COM может "успешно" создать
геометрически бессмысленную операцию.

Health-статус фичи читается так: `feature.HealthStatus` — значения см.
`tools/enum_lookup.py HealthStatusEnum` (например, `kUpToDateHealth=11778`,
`kDriverLostHealth=11780` — именно этот статус был у нерабочего отверстия).

## Сборки — сопряжения (constraints), живой тест

**Подтверждено рабочим тестом** — два экземпляра детали стянуты «лицом к лицу» через `AddMateConstraint`,
проверено рендером (грани сошлись вплотную).

```python
def find_face(occ, target_area):
    for f in occ.SurfaceBodies.Item(1).Faces:
        if abs(f.Evaluator.Area - target_area) < 0.5:
            return f

f1 = find_face(occ1, 60.0)
f2 = find_face(occ2, 60.0)
proxy1 = occ1.CreateGeometryProxy(f1)   # грань детали → грань "в контексте сборки"
proxy2 = occ2.CreateGeometryProxy(f2)
compDef.Constraints.AddMateConstraint(proxy1, proxy2, 0)   # 0 = зазор (offset), см
```

`CreateGeometryProxy(entity)` вызывается с ОДНИМ аргументом на чистом dynamic dispatch — второй
(`ByRef`) параметр из IDL можно не передавать, pywin32 отдаёт его как обычный возврат функции.

## Спецификация (BOM)

**Подтверждено рабочим тестом** — прочитана спецификация сборки, обозначение (Part Number) из
iProperties детали корректно попало в строку BOM.

```python
bom = compDef.BOM
bom.StructuredViewEnabled = True
for i in range(1, bom.BOMViews.Count + 1):          # Item(enum) по значению не работает —
    v = bom.BOMViews.Item(i)                          # перебирать по индексу и сверять .ViewType
    if v.ViewType == 62466:                            # kStructuredBOMViewType
        view = v
for row in view.BOMRows:
    row.ItemNumber
    row.TotalQuantity
    comp = row.ComponentDefinitions.Item(1)
    comp.Document.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
```
Результат теста: `item=1 qty=2 partnumber=1000.0002.0004 name=TestPart.ipt` — ровно два
экземпляра детали, обозначение подтянулось автоматически.

## Массив (Rectangular Pattern)

**Подтверждено рабочим тестом** — с ещё одним примером «тихого» отказа корректности.

```python
parentColl = app.TransientObjects.CreateObjectCollection()
parentColl.Add(holeFeat)                                    # что размножаем — любая фича

pattern = compDef.Features.RectangularPatternFeatures.Add(
    parentColl,
    xEdge, True, 3, 2.5, kDefault, None,                     # направление X: ребро, count, spacing(см)
    None, True, 1, 0, kDefault, None,                        # направление Y — отключено (count=1)
    kIdenticalCompute, kIdentical)                            # 47361, 33793
```

### ⚠️ Грабли: копия массива на границе детали тихо обрезается

При `XCount=3, XSpacing=2.5` одна из копий отверстия попала прямо на край пластины и вышла не
круглым отверстием, а надрезом с края (подтверждено рендером). Feature создаётся без ошибки,
`HealthStatus` в норме — узнать о проблеме можно только визуально или сравнением объёма/числа
граней. **Правило то же, что и с отверстием**: после массива всегда проверять фактическую
геометрию, не доверять одному факту отсутствия исключения.

## Проверка пересечений в сборке (Interference)

**Подтверждено рабочим тестом** — оба случая: отсутствие пересечения и реальное пересечение с
точным объёмом.

```python
set1 = app.TransientObjects.CreateObjectCollection(); set1.Add(occ1)
set2 = app.TransientObjects.CreateObjectCollection(); set2.Add(occ2)
result = compDef.AnalyzeInterference(set1, set2)
result.Count                        # 0 — нет пересечений
result.Item(i).Volume               # объём пересечения в см³, если есть
```

Проверено на двух сценариях: детали состыкованы впритык (`Count=0`, верно — касание не
пересечение) и детали намеренно сдвинуты внахлёст (`Count=1`, `Volume=40.93` см³ — совпало с
расчётной геометрией перекрытия).

### ⚠️ Грабли: `occurrence.SetTransformation(...)` не существует — это свойство, не метод

Правильно: `occurrence.Transformation = matrix` (присваивание свойству), а не вызов метода.
Ошибка вида `AttributeError: ... Did you mean: 'Transformation'?` — верный сигнал именно об этом.
Также: если на occurrence есть активное сопряжение (constraint), прямое присваивание
`Transformation` может не дать реального сдвига (решатель сборки пересчитывает позицию по
ограничениям) — сначала удалить/подавить связанные constraints, если нужно двигать деталь вручную.

## ⚠️ Грабли: зависшие msedgewebview2.exe ломают повторный запуск пульта

При частых перезапусках `webview_app.py` в разработке (`TaskStop` на bash-обёртке python) дерево
дочерних процессов WebView2 (`msedgewebview2.exe` — браузерный, GPU, рендер-процессы) не всегда
завершается вместе с родителем. За несколько перезапусков накопилось ~40 таких процессов, и
следующее открытое окно пульта отрисовывалось битым — 160×28 пикселей вместо заданных 820×680,
пустое. Никакой ошибки при этом Python не бросал.

**Диагностика**: `tasklist /FI "IMAGENAME eq msedgewebview2.exe"` — если процессов десятки, это
оно.

**Фикс**: `taskkill /F /IM msedgewebview2.exe` + `taskkill /F /IM python.exe`, затем запустить
пульт заново.

**Для собранного `.exe` через PyInstaller это менее критично** — пользователь обычно не
перезапускает приложение по 10 раз за минуту, как это происходит при разработке.

## Источники (внешнее исследование, 2026)

Для ориентира, что уже делают в индустрии на стыке Inventor + AI:

- [ipt-mcp](https://github.com/bimwright/ipt-mcp) — open-source MCP-сервер для Inventor 2022–2027,
  58 инструментов/13 категорий (query, document, parameters, properties, sketch, feature, export,
  assembly, assembly_query, code-escape-hatch, "ToolBaker" — сохранение рабочих сценариев как
  переиспользуемых именованных инструментов). Архитектурно тяжелее нашего (отдельный .NET add-in
  + именованный канал), но набор возможностей — ориентир для роадмапа. Они переводят все длины в
  мм на границе API (Inventor изнутри — см) — стоит перенять для `core.py`.
- [Autodesk Assistant в Inventor 2027](https://blog.autodesk.io/from-commands-to-conversations-exploring-autodesk-assistant-in-inventor-2027/)
  (технический preview) — умеет только анализировать модель на естественном языке (иерархия
  сборки, масса, отсутствие материала), **не может** создавать/менять геометрию — только
  подсказывает, что делать руками. Наш мост уже мощнее в этом смысле.
- Autodesk Design Copilot (2026), Ansys AI+, Inventor Shape Generator, nTopology — соседние
  направления (предсказание следующего шага, топологическая оптимизация, симуляция) — не
  пересекаются с задачей «пульт с кнопками», но полезно знать контекст рынка.

## Скриншоты для проверки — какой способ использовать

Пробовали два способа. **Рекомендуется первый**:

1. **`ActiveView.SaveAsBitmap(path, width, height)`** — рендерит сцену средствами самого Inventor
   напрямую в файл. Не зависит от того, свёрнуто ли окно, что на переднем плане, не требует
   поиска HWND. Единственный минус — отдаёт `.bmp`, для показа нужно сконвертировать в `.png`
   (через `System.Drawing` в PowerShell: `[System.Drawing.Image]::FromFile(...).Save(..., Png)`).
2. Захват экрана через `System.Windows.Forms.Screen` + `CopyFromScreen` — **осторожно**: если
   окно Inventor не на переднем плане или не найдено по заголовку, снимается **весь рабочий
   стол**, включая другие открытые приложения пользователя (был случай — попал чужой чат другого
   инструмента). Если такой способ всё же нужен, сначала находить окно через `PrintWindow` по
   HWND конкретного процесса, никогда не снимать весь экран вслепую.
