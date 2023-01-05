**PivotItem**



代表数据透视表字段中的项目。

**说明**

这些项目是某个字段类型中的各个数据项。**PivotItem** 对象是 **PivotItems** 集合的成员。**PivotItems** 集合包含某个 **PivotField** 对象中所有的项目。使用 **PivotItems**( *index*)（其中 *index* 是项目索引号或名称）可返回一个 **PivotItem** 对象。下例隐藏“Sheet3”上第一张数据透视表中“Year”字段中包含“1998”的所有数据项。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("sheet3").PivotTables(1).PivotFields("year").PivotItems("1998").Visible = false` |

**方法**

|                                                              | 名称        | 说明                                                         |
| ------------------------------------------------------------ | ----------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**  | table { word-break:break-all; }删除对象。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DrillTo** | table { word-break:break-all; }**DrillTo** 方法支持从 PivotItem 中深化到指定的透视字段。 |

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**        | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**            | table { word-break:break-all; }返回一个 **String** 值，它代表数据透视项的标签文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChildItems**         | 返回一个对象，它代表一个数据透视表数据项（**PivotItem** 对象）或所有数据项的集合（**PivotItems** 对象），该集合可以由指定字段中的子项组成，也可以由指定数据项中的子项组成。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**            | table { word-break:break-all; }返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataRange**          | 返回一个 **Range** 对象，如下表所示。只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DrilledDown**        | 如果指定数据透视表字段或数据透视表项的标志设置为“drilled”（展开或可见），则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Formula**            | table { word-break:break-all; }返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsCalculated**       | table { word-break:break-all; }如果数据透视表字段或数据透视表项为计算字段或计算项，则此属性为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LabelRange**         | table { word-break:break-all; }返回一个 **Range** 对象，它表示数据透视表中所有包含数据项的单元格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**               | table { word-break:break-all; }返回或设置一个 **String** 值，它代表对象的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**             | table { word-break:break-all; }返回指定对象的父对象。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentItem**         | 返回一个 **PivotItem** 对象，该对象表示父 **PivotField** 对象中的父数据透视表项（必须对该字段分组使其有一个父级）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentShowDetail**   | 如果指定数据项是由于其某个父数据项正在显示明细数据而进行显示的，则该值为 **True**。如果由于其父数据项隐藏明细数据而使该数据项未显示，则该值为 **False**。本属性仅对已分组的数据项有效。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**           | 返回或设置一个 **Long** 值，它代表当前正在显示数据项时，该数据项字段中数据项的位置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RecordCount**        | 返回数据透视表缓存中的记录数目或包含指定内容的缓存记录数目。**Long** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowDetail**         | table { word-break:break-all; }如果扩展了指定区域的分级显示（从而行或列的明细数据可见），则为 **True**。指定区域必须为分级显示中的单个汇总列或汇总行。**Variant** 型，可读写。对于 **PivotItem** 对象（如果该区域在数据透视表中，则为 **Range** 对象），当数据项显示明细数据时，此属性设为 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceName**         | 返回一个 **Variant** 值，它代表指定的对象出现在指定的数据透视表的原始源数据中时的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceNameStandard** | 返回一个 **String** 类型的数值，该数值表示以标准英文（美国）格式设置的数据透视表项目的源名称。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StandardFormula**    | 返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**              | table { word-break:break-all; }返回或设置一个 **String** 值，它代表数据透视表中指定的数据项的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**            | table { word-break:break-all; }返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。 |

**成员方法**

#### **PivotItem.Delete**

table { word-break:break-all; }

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **PivotItem** 对象的变量。

#### **PivotItem.DrillTo**

table { word-break:break-all; }

**DrillTo** 方法支持从 PivotItem 中深化到指定的透视字段。

**语法**

**express.DrillTo(PivotItemName)**

*express*   一个代表 **PivotItem** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                      |
| --------------- | ------------- | ------------ | ----------------------------- |
| *PivotItemName* | 必选          | **String**   | 要深化到的 PivotItem 的名称。 |

**说明**

table { word-break:break-all; }

对于 OLAP 数据源，要深化到的透视字段必须与正在深化的 PivotItem 位于同一层次结构中，或者，如果多个属性层次结构彼此相邻地放置在行或列上，则要深化到的透视字段必须是彼此相邻的属性层次结构之一；不可以将用户层次结构放置在正在深化的 PivotItem 的透视字段与要深化到的透视字段之间。如果不符合这些条件，则会返回一个运行时错误。

**成员属性**

#### **PivotItem.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotItem** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **PivotItem.Caption**

table { word-break:break-all; }

返回一个 **String** 值，它代表数据透视项的标签文本。

**语法**

**express.Caption**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

table { word-break:break-all; }

下表显示 **Caption** 属性及其相关属性的示例值，假设存在一个具有唯一名称为“[Europe].[France].[Paris]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，并且存在一个包含名为“Paris”项的非 OLAP 数据源。

| 属性           | 值（OLAP 数据源）                 | 值（非 OLAP 数据源）        |
| -------------- | --------------------------------- | --------------------------- |
| **Caption**    | Paris                             | Paris                       |
| **Name**       | [Europe].[France].[Paris]（只读） | Paris                       |
| **SourceName** | [Europe].[France].[Paris]（只读） | （与 SQL 属性值相同；只读） |
| **Value**      | [Europe].[France].[Paris]（只读） | Paris                       |

在将一个索引值指定到 **PivotItems** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源）                                | 语法（非 OLAP 数据源）         |
| -------------------------------------------------- | ------------------------------ |
| expression.PivotItems("[Europe].[France].[Paris]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源）       | 名称（非 OLAP 数据源） |
| ------------------------- | ---------------------- |
| [Europe].[France].[Paris] | Paris                  |

#### **PivotItem.ChildItems**

返回一个对象，它代表一个数据透视表数据项（**PivotItem** 对象）或所有数据项的集合（**PivotItems** 对象），该集合可以由指定字段中的子项组成，也可以由指定数据项中的子项组成。只读。

**语法**

**express.ChildItems**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

此属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

**示例**

此示例将“vegetables”数据项的所有子数据项添加到一张新工作表上的数据清单中。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let nwSheet = Worksheets.Add() nwSheet.Activate() let pvtTable = Worksheets.Item("Sheet2").Range("A1").PivotTable let rw = 0 for(let i=1;i <= pvtTable.PivotFields("product").PivotItems("vegetables").**ChildItems**.Count;i++) {     rw++     nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("product").PivotItems("vegetables").ChildItems.Item(i).Name }` |

#### **PivotItem.Creator**

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotItem.DataRange**

返回一个 **Range** 对象，如下表所示。只读。

**语法**

**express.DataRange**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

| 对象             | 数据区域               |
| ---------------- | ---------------------- |
| “数据”字段       | 字段中包含的数据       |
| 行、列或者页字段 | 字段中的数据项         |
| 数据项           | 数据项规范所限定的数据 |

**示例**

此示例选定名为“REGION”字段中的数据透视表数据项。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable Worksheets.Item("Sheet1").Activate() pvtTable.PivotFields("REGION").**DataRange**.Select()` |

#### **PivotItem.DrilledDown**

如果指定数据透视表字段或数据透视表项的标志设置为“drilled”（展开或可见），则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DrilledDown**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

此属性仅对 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源有效。

如果字段或项被隐藏，则不可设置此属性。

**示例**

此示例将活动工作表上第三个数据透视表中状态字段的所有项的标记设置为“not drilled”。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `ActiveSheet.PivotTables("PivotTable3").PivotFields("state").**DrilledDown** = false` |

#### **PivotItem.Formula**

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

**语法**

**express.Formula**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

table { word-break:break-all; }

此属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回空字符串。如果单元格包含公式，**Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

#### **PivotItem.IsCalculated**

table { word-break:break-all; }

如果数据透视表字段或数据透视表项为计算字段或计算项，则此属性为 **True**。**Boolean** 类型，只读。

**语法**

**express.IsCalculated**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

table { word-break:break-all; }

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性始终返回 **False**。

#### **PivotItem.LabelRange**

table { word-break:break-all; }

返回一个 **Range** 对象，它表示数据透视表中所有包含数据项的单元格。只读。

**语法**

**express.LabelRange**

*express*   一个代表 **PivotItem** 对象的变量。

#### **PivotItem.Name**

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

table { word-break:break-all; }

下表显示了 **Name** 属性和相关属性的示例值，同时假设有一个具有唯一名称为“[Europe].[France].[Paris]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，有一个包含项名称为“Paris”的非 OLAP 数据源。

| 属性           | 值（OLAP 数据源）                 | 值（非 OLAP 数据源）        |
| -------------- | --------------------------------- | --------------------------- |
| **Caption**    | Paris                             | Paris                       |
| **Name**       | [Europe].[France].[Paris]（只读） | Paris                       |
| **SourceName** | [Europe].[France].[Paris]（只读） | （与 SQL 属性值相同，只读） |
| **Value**      | [Europe].[France].[Paris]（只读） | Paris                       |

在将一个索引值指定到 **PivotItems** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源）                                | 语法（非 OLAP 数据源）         |
| -------------------------------------------------- | ------------------------------ |
| expression.PivotItems("[Europe].[France].[Paris]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源）       | 名称（非 OLAP 数据源） |
| ------------------------- | ---------------------- |
| [Europe].[France].[Paris] | Paris                  |

#### **PivotItem.Parent**

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotItem** 对象的变量。

#### **PivotItem.ParentItem**

返回一个 **PivotItem** 对象，该对象表示父 **PivotField** 对象中的父数据透视表项（必须对该字段分组使其有一个父级）。只读。

**语法**

**express.ParentItem**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

该属性不可用于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

**示例**

本示例显示包含活动单元格的项的父项名称。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").Activate() MsgBox("This item is a subitem of " + ActiveCell.PivotItem.**ParentItem**.Name)` |

#### **PivotItem.ParentShowDetail**

如果指定数据项是由于其某个父数据项正在显示明细数据而进行显示的，则该值为 **True**。如果由于其父数据项隐藏明细数据而使该数据项未显示，则该值为 **False**。本属性仅对已分组的数据项有效。**Boolean** 类型，只读。

**语法**

**express.ParentShowDetail**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

该属性不可用于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

**示例**

如果包含活动单元格的项由于其父项正在显示明细数据而可见，则本示例将显示一条消息。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").Activate() let pvtItem = ActiveCell.PivotItem if(pvtItem.**ParentShowDetail** == true) {     MsgBox("Parent item is showing detail") }` |

#### **PivotItem.Position**

返回或设置一个 **Long** 值，它代表当前正在显示数据项时，该数据项字段中数据项的位置。

**语法**

**express.Position**

*express*   一个代表 **PivotItem** 对象的变量。

**示例**

此示例显示包含活动单元格的数据透视表项的位置数字。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").Activate() MsgBox("The active item is in position number " + ActiveCell.PivotItem.**Position**)` |

#### **PivotItem.RecordCount**

返回数据透视表缓存中的记录数目或包含指定内容的缓存记录数目。**Long** 型，只读。

**语法**

**express.RecordCount**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

此属性反映了查询高速缓存时的瞬时状态。该高速缓存可对不同查询而更改。

**示例**

此示例显示“Products”字段中包含“Kiwi”的缓存记录数目。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `MsgBox(Worksheets.Item(1).PivotTables("Pivot1").PivotFields("Product").PivotItems("Kiwi").**RecordCount**)` |

#### **PivotItem.ShowDetail**

table { word-break:break-all; }

如果扩展了指定区域的分级显示（从而行或列的明细数据可见），则为 **True**。指定区域必须为分级显示中的单个汇总列或汇总行。**Variant** 型，可读写。对于 **PivotItem** 对象（如果该区域在数据透视表中，则为 **Range** 对象），当数据项显示明细数据时，此属性设为 **True**。

**语法**

**express.ShowDetail**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

table { word-break:break-all; }

此属性不可用于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

如果指定区域不在数据透视表中，则下列声明为真：

- 指定区域必须在单个汇总行或汇总列中。
- 如果指定行或列的任意子级处于隐藏状态，则此属性返回 **False**。
- 将此属性设置为 **True** 相当于显示指定汇总行或汇总列的所有子级内容。
- 将此属性设置为 **False** 相当于隐藏汇总行或汇总列的所有子级内容。

如果指定区域为数据透视表，则当该区域连续时，就可以同时对多个单元格设置此属性。仅当指定区域为单个单元格时，才能返回此属性的值。

#### **PivotItem.SourceName**

返回一个 **Variant** 值，它代表指定的对象出现在指定的数据透视表的原始源数据中时的名称。

**语法**

**express.SourceName**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

如果用户在创建数据透视表之后重命名数据项，此属性的值可能会与当前数据项名称不同。

下表显示 **SourceName** 属性及其相关属性的示例值，假设存在唯一名称为“[Europe].[France].[Paris]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，并且存在包含名为“Paris”的项的非 OLAP 数据源。

| 属性           | 值（OLAP 数据源）                 | 值（非 OLAP 数据源）        |
| -------------- | --------------------------------- | --------------------------- |
| **Caption**    | Paris                             | Paris                       |
| **Name**       | [Europe].[France].[Paris]（只读） | Paris                       |
| **SourceName** | [Europe].[France].[Paris]（只读） | （与 SQL 属性值相同，只读） |
| **Value**      | [Europe].[France].[Paris]（只读） | Paris                       |

在将一个索引值指定到 **PivotItems** 集合中时，您可以使用下表所示的语法。

| 语法（OLAP 数据源）                                | 语法（非 OLAP 数据源）         |
| -------------------------------------------------- | ------------------------------ |
| expression.PivotItems("[Europe].[France].[Paris]") | expression.PivotItems("Paris") |

在使用 **Item** 属性引用集合中的特定成员时，可以使用下表所示的文本索引名称。

| 名称（OLAP 数据源）       | 名称（非 OLAP 数据源） |
| ------------------------- | ---------------------- |
| [Europe].[France].[Paris] | Paris                  |

**示例**

本示例显示包含活动单元格的数据项的原始名称（在源数据库中的名称）。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").Activate() ActiveSheet.PivotTables(1).PivotSelect("1998", xlDataAndLabel) MsgBox("The original item name is " + ActiveCell.PivotItem.**SourceName**)` |

#### **PivotItem.SourceNameStandard**

返回一个 **String** 类型的数值，该数值表示以标准英文（美国）格式设置的数据透视表项目的源名称。只读。

**语法**

**express.SourceNameStandard**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

当某一项具有本地化版本，并且它的 **SourceNameStandard** 属性值不同于 **SourceName** 属性值（例如带日期格式）时，使用该属性。

**示例**

本示例显示活动数据透视表的第五个字段上第六个项目的源名称。本示例假定数据透视表位于活动工作表上，并且数据源至少包括五个字段，且每个字段至少包括六个项目。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function CheckSourceNameStandard() {     let pvtTable = ActiveSheet.PivotTables(1)     let pvtField = pvtTable.PivotFields(5)     let pvtItem = pvtField.PivotItems(6)     // Display source name.     MsgBox("The source name is: " + pvtItem.**SourceNameStandard**) }` |

#### **PivotItem.StandardFormula**

返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。

**语法**

**express.StandardFormula**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

**StandardFormula** 属性主要影响具有日期或数字格式的项目名称。该属性提供了一种方法可指定或查询给定计算项的公式。

**StandardFormula** 属性是国际通用的，**Formula** 属性则不是。

**示例**

本示例向“Decimals”字段中添加 10，并将其显示为数据字段中的计算项。本示例假定数据透视表位于活动工作簿上，并且标题为“Decimals”的字段位于模拟运算表中。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function UseStandardFomula() {     let pvtTable = ActiveSheet.PivotTables(1)     // Change calculated field of decimals by adding '10'.     pvtTable.CalculatedFields().Item(1).**StandardFormula** = "Decimals + 10" }` |

#### **PivotItem.Value**

table { word-break:break-all; }

返回或设置一个 **String** 值，它代表数据透视表中指定的数据项的名称。

**语法**

**express.Value**

*express*   一个代表 **PivotItem** 对象的变量。

#### **PivotItem.Visible**

table { word-break:break-all; }

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **PivotItem** 对象的变量。

**说明**

table { word-break:break-all; }

如果数据透视表中的某一项在当前是可见的，则该项的 **Visible** 属性为 **True**。

适用环境：web

适用平台：windows/linux