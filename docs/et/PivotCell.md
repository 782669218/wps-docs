**PivotCell**



代表数据透视表中的一个单元格。

**说明**

使用 **Range** 集合的 **PivotCell** 属性可返回一个 **PivotCell** 对象。

返回 **PivotCell** 对象后，可以使用 **ColumnItems** 或 **RowItems** 属性来确定 **PivotItems** 集合，对应于代表所选编号的列轴或行轴上的项目。下例使用 **PivotCell** 对象的 **ColumnItems** 属性来返回一个 **PivotItemList** 集合。

返回了 **PivotCell** 对象后，可以使用 **PivotCellType** 属性来确定某个特定区域是什么单元格类型。下例确定数据透视表中的单元格 A5 是否是数据项目并通知用户。此示例假定活动工作表上存在数据透视表而且单元格 A5 包含在该数据透视表中。如果单元格 A5 不在数据透视表中，该示例会处理运行时错误。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/**/ function test(){ try {         // Determine if cell A5 is a data item in the PivotTable.         if(Application.Range("A5").PivotCell.PivotCellType == xlPivotCellValue) {             alert ("The PivotCell at A5 is a data item.")         }         else {             alert ("The PivotCell at A5 is a data item.")         }     } 									     catch(exception){         alert("The chosen cell is not in a PivotTable.")     } }` |

此示例确定单元格 B5 的数据项所在的列字段。然后判断列字段标题是否与“Inventory”相匹配，并通知用户。此示例假定数据透视表位于活动工作表上，并且工作表的 B 列包含数据透视表的列字段。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/**/ function test(){   // Determine if there is a match between the item and column field.     if(Application.Range("B5").PivotCell.ColumnItems.Item(1) == "Inventory") {         alert ("Item in B5 is a member of the 'Inventory' column field.")     }     else {         alert ("Item in B5 is not a member of the 'Inventory' column field.")     } }` |

**方法**

|                                                              | 名称              | 说明                                 |
| ------------------------------------------------------------ | ----------------- | ------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DiscardChange** | 放弃对数据透视表中指定单元格的更改。 |

**属性**

|                                                              | 名称                       | 说明                                                         |
| ------------------------------------------------------------ | -------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**            | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CellChanged**            | 返回自创建数据透视表以来，或上次执行提交操作以来，数据透视表值单元格是否经过了编辑或重新计算。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColumnItems**            | 返回一个 **PivotItemList** 集合，该集合对应于表示选定区域的列轴上的项目。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomSubtotalFunction** | 返回 **PivotCell** 对象的自定义分类汇总函数字段的设置。**XlConsolidationFunction** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataField**              | 返回一个 **PivotField** 对象，该对象与所选数据字段相对应。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataSourceValue**        | 为数据透视表中经过编辑的单元格，返回上次从数据源检索到的值。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MDX**                    | 返回一个元组，该元组提供了使用 OLAP 数据源的数据透视表中指定值单元格的完整 MDX 坐标。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                 | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotCellType**          | 返回一个 **XlPivotCellType** 常量，该常量标识单元格所对应的数据透视表实体。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotColumnLine**        | 在列上返回特定 **PivotCell** 对象的 **PivotLine**。只读 **PivotLine** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotField**             | 返回一个 **PivotField** 对象，它代表指定区域左上角所在的数据透视表字段。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotItem**              | 返回一个 **PivotItem** 对象，它代表指定区域左上角所在的数据透视表项。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotRowLine**           | 在行上返回特定 **PivotCell** 对象的 PivotLine。只读 **PivotLine** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotTable**             | 返回一个 **PivotTable** 对象，它代表与 PivotCell 相关联的数据透视表。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**                  | 返回一个 **Range** 对象，它代表指定的 PivotCell 所应用的区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RowItems**               | 返回一个 **PivotItemList** 集合，该集合对应于分类轴上表示选定单元格的项目。 |

**成员方法**

#### **PivotCell.DiscardChange**

放弃对数据透视表中指定单元格的更改。

**语法**

**express.DiscardChange()**

*express*   一个代表 **PivotCell** 对象的变量。

**返回值**

Nothing

**说明**

对于基于 OLAP 数据源的数据透视表中的单元格，**DiscardChange** 方法删除在该指定单元格中输入的值或公式，然后运行数据透视表更新从数据源检索最新值。该指定单元格的对应数据源值将由此方法设置为 **NULL**。该方法还将从 ET 更改列表中删除对此单元格所做的所有更改，以使这些更改不再包含在对该数据源的 **UPDATE CUBE** 语句中。此方法还将重新创建在事务中发生的所有其他更改的会话状态。这包括执行 **ROLLBACK TRANSACTION** 语句，然后执行 **UPDATE CUBE** 语句应用除该指定单元格之外的所有更改。对于基于非 OLAP 数据源的数据透视表中的单元格，此方法清除编辑过的单元格。

**成员属性**

#### **PivotCell.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotCell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **PivotCell.CellChanged**

返回自创建数据透视表以来，或上次执行提交操作以来，数据透视表值单元格是否经过了编辑或重新计算。只读

**语法**

**express.CellChanged**

*express*   一个代表 **PivotCell** 对象的变量。

**说明**

table { word-break:break-all; }

**CellChanged** 属性的值默认为 **xlCellNotChanged** 。

对于使用非 OLAP 数据源的数据透视表，此属性的值只能为 **xlCellNotChanged** 或 **xlCellChanged**。对于未编辑的单元格，其值为 **xlCellNotChanged**；对于编辑过的单元格，其值为 **xlCellChanged**。放弃更改将把其值设置为 **xlCellNotChanged**。

应用并保存更改只适用于使用 OLAP 数据源的数据透视表。下面的列表是 **CellChange** 属性各种可能的状态的说明，它们中适用于使用 OLAP 数据源的数据透视表：

- **xlCellNotChanged** - 自创建数据透视表以来，或者自上次执行保存或放弃更改操作以来，单元格尚未编辑或重新计算过（如果单元格包含公式）。
- **xlCellChanged** - 自创建数据透视表以来，或者自上次执行应用更改或保存更改操作以来，单元格已经过编辑或重新计算，但是尚未应用该更改（尚未对该更改运行 **UPDATE CUBE** 语句）。
- **xlCellChangeApplied** - 自创建数据透视表以来，或者自上次执行应用更改、保存更改或放弃更改操作以来，单元格已经过编辑或重新计算，并且已经应用了该更改（已经对该更改运行 **UPDATE CUBE** 语句）。

下表说明不同的用户操作如何影响使用 OLAP 数据源的数据透视表中 **CellChanged** 属性的设置。

| 用户操作                                                     | 无公式单元格的 CellChanged 属性的设置                        | 有公式单元格的 CellChanged 属性的设置                        |
| ------------------------------------------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------ |
| 在一个或多个单元格中输入值或公式。                           | 对这些单元格，设置为 **xlCellChanged**。                     | 对这些单元格，设置为 **xlCellChanged**。                     |
| 重新计算含公式的一个或多个单元格（手动计算 (F9)，或由 ET 自动计算）。 | 不适用                                                       | 对这些单元格，设置为 **xlCellChanged**。                     |
| 保存（提交）更改。                                           | 对所有经过编辑的无公式单元格，设置为 **xlCellNotChanged**。  | 对所有经过编辑的有公式单元格，设置为 **xlCellChangeApplied**。 |
| 放弃所有更改。                                               | 对所有经过编辑的无公式单元格，设置为 **xlCellNotChanged**。  | 对所有经过编辑的有公式单元格，设置为 **xlCellNotChanged**。  |
| 放弃单个单元格中的更改。                                     | 只对该单元格设置为 **xlCellNotChanged**。                    | 只对该单元格设置为 **xlCellNotChanged**。                    |
| 一次操作清除多个单元格。                                     | 对这些单元格设置为 **xlCellNotChanged**。                    | 对这些单元格设置为 **xlCellNotChanged**。                    |
| 清除一个单元格。                                             | 只对该单元格设置为 **xlCellNotChanged**。                    | 只对该单元格设置为 **xlCellNotChanged**。                    |
| 应用值之前，执行撤消操作，将该值改回先前编辑的值。           | 对所有经过编辑的无公式单元格，设置为 **xlCellChanged**。     | 对所有经过编辑的有公式单元格，设置为 **xlCellChanged**。     |
| 应用值之后，执行撤消操作，将该值改回先前编辑的值。           | 对所有经过编辑的无公式单元格，设置为 **xlCellChangedApplied**。 | 对所有经过编辑的有公式单元格，设置为 **xlCellChangeApplied**。 |

#### **PivotCell.ColumnItems**

返回一个 **PivotItemList** 集合，该集合对应于表示选定区域的列轴上的项目。

**语法**

**express.ColumnItems**

*express*   一个代表 **PivotCell** 对象的变量。

**示例**

本示例判断单元格 B5 中的数据项是否在第一个列字段中的“Inventory”项目之下，并通知用户。本示例假定数据透视表位于活动工作表上，并且第 B 列包含数据透视表的列字段。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ // Determine if there is a match between the item and column field. if(Application.Range("B5").PivotCell.ColumnItems.Item(1) == "Inventory") {     MsgBox("Item in B5 is a member of the 'Inventory' column field.") } else {     MsgBox("Item in B5 is not a member of the 'Inventory' column field.") } }` |

#### **PivotCell.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotCell** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotCell.CustomSubtotalFunction**

返回 **PivotCell** 对象的自定义分类汇总函数字段的设置。**XlConsolidationFunction** 类型，只读。

**语法**

**express.CustomSubtotalFunction**

*express*   一个代表 **PivotCell** 对象的变量。

**说明**

| **XlConsolidationFunction** 可为以下 **XlConsolidationFunction** 常量之一。 |
| ------------------------------------------------------------ |
| **xlAverage**                                                |
| **xlCount**                                                  |
| **xlCountNums**                                              |
| **xlMax**                                                    |
| **xlMin**                                                    |
| **xlProduct**                                                |
| **xlStDev**                                                  |
| **xlStDevP**                                                 |
| **xlSum**                                                    |
| **xlUnknown**                                                |
| **xlVar**                                                    |
| **xlVarP**                                                   |

如果 **PivotCell** 对象类型不是自定义分类汇总，那么 **CustomSubtotalFunction** 属性将返回一个错误。该属性只应用于非 OLAP 源数据?（非 OLAP 源数据：非 OLAP 源数据是数据透视表或数据透视图使用的基本数据，该数据来自 OLAP 数据库之外的源。这些数据源包括关系数据库、 ET 工作表中的清单以及文本文件数据库。）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定单元格 C20 是否包含使用计数合并函数的自定义分类汇总函数，并通知用户。本示例假定数据透视表位于活动工作表上。*/ function test(){   try{         // Determine if custom subtotal function is a count function.         if(Application.Range("C20").PivotCell.CustomSubtotalFunction == xlCount) {             MsgBox("The custom subtotal function is a Count.")         }         else             alert("The custom subtotal function is not a Count.")         } 									     catch(exception) {         alert("The selected cell is not a custom subtotal function.")     } }` |

#### **PivotCell.DataField**

返回一个 **PivotField** 对象，该对象与所选数据字段相对应。

**语法**

**express.DataField**

*express*   一个代表 **PivotCell** 对象的变量。

**说明**

如果 **PivotCell** 对象不是 **XlPivotCellType** 允许的常量：**xlPivotCellTypeDataField**、**xlPivotCellTypeSubtotal** 或 **xlPivotCellTypeGrandTotal** 之一，那么该属性将返回一个错误。

**示例**

本示例确定单元格 L10 是否在数据透视表的数据字段中，然后通过通知用户来返回对应于该单元格的数据透视表字段，或处理运行错误。本示例假定数据透视表位于活动工作表上。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   try{         alert(Application.Range("L10").PivotCell.DataField)     } 									     catch(exception) {         alert("The selected range is not in the data field of the PivotTable.")     } }` |

#### **PivotCell.DataSourceValue**

为数据透视表中经过编辑的单元格，返回上次从数据源检索到的值。只读

**语法**

**express.DataSourceValue**

*express*   一个代表 **PivotCell** 对象的变量。

**说明**

每次编辑数据透视表值区域中的某个单元格时，在编辑发生之前，**DataSourceValue** 属性将保存上次从数据源检索到的值。对于尚未编辑的数据透视表值单元格，或者尚未显式检索数据源值的数据透视表值单元格，此属性将返回 **NULL**。对于使用 OLAP 数据源的数据透视表，**DataSourceValue** 属性的值将从单独的连接检索，以确保其不包含用户可能已执行的任何回写操作的值。

如果对数据透视表值区域之外的单元格读取其 **DataSourceValue** 属性，那么会生成运行时错误。

#### **PivotCell.MDX**

返回一个元组，该元组提供了使用 OLAP 数据源的数据透视表中指定值单元格的完整 MDX 坐标。只读

**语法**

**express.MDX**

*express*   一个代表 **PivotCell** 对象的变量。

**说明**

元组中由 **MDX** 属性返回的维度包含行坐标和列坐标，以及报表筛选坐标。对于数据透视表值区域之外的单元格，以及数据透视表之外的单元格，访问此属性将生成运行时错误。对于在报表筛选字段中有多项选择的数据透视表，访问此属性也将生成运行时错误。

#### **PivotCell.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotCell** 对象的变量。

#### **PivotCell.PivotCellType**

返回一个 **XlPivotCellType** 常量，该常量标识单元格所对应的数据透视表实体。只读。

**语法**

**express.PivotCellType**

*express*   一个代表 **PivotCell** 对象的变量。

**说明**

| **XlPivotCellType** 可为以下 **XlPivotCellType** 常量之一。  |
| ------------------------------------------------------------ |
| **xlPivotCellBlankCell?**数据透视表中的一个结构空白单元格。  |
| **xlPivotCellCustomSubtotal** 自定义分类汇总的行或列区域中的一个单元格。 |
| **xlPivotCellDataField?**一个数据字段标签（不是“数据”按钮）。 |
| **xlPivotCellDataPivotField**“数据”按钮。                    |
| **xlPivotCellGrandTotal?**总计行或列区域中的一个单元格。     |
| **xlPivotCellPageFieldItem?**显示页字段的选定项的单元格。    |
| **xlPivotCellPivotField?**字段的按钮（不是“数据”按钮）。     |
| **xlPivotCellPivotItem?**不是分类汇总、总计、自定义分类汇总或空行的行或列区域中的单元格。 |
| **xlPivotCellSubtotal?**分类汇总的行或列区域中的单元格。     |
| **xlPivotCellValue?**数据区区域（除空行以外）中的任一单元格。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定数据透视表中的单元格 A5 是否是一个数据项，并通知用户。本示例假定数据透视表位于活动工作表上。如果单元格 A5 不在数据透视表中，则本示例处理运行错误。*/ function test(){  try{         // Determine if cell A5 is a data item in the PivotTable.         if(Application.Range("A5").PivotCell.PivotCellType == xlPivotCellValue) {             alert ("The cell at A5 is a data item.")         }         else {         alert("The cell at A5 is not a data item.")         }     }     catch(exception) {         alert ("The chosen cell is not in a PivotTable." )     } }` |

#### **PivotCell.PivotColumnLine**

在列上返回特定 **PivotCell** 对象的 **PivotLine**。只读 **PivotLine** 类型。

**语法**

**express.PivotColumnLine**

*express*   一个代表 **PivotCell** 对象的变量。

**说明**

如果 PivotCell 在行上，则 **PivotColumnLine** 属性返回一个运行时错误。

如果 PivotCell 在列上，则 **PivotColumnLine** 属性返回列 **PivotLine** 对象。

如果 PivotCell 在数据区域中，则 **PivotColumnLine** 属性返回对应的列 **PivotLine** 对象。

#### **PivotCell.PivotField**

返回一个 **PivotField** 对象，它代表指定区域左上角所在的数据透视表字段。

**语法**

**express.PivotField**

*express*   一个代表 **PivotCell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示包含活动单元格的数据透视表字段的名称。*/ function test(){ Worksheets.Item("Sheet1").Activate() alert("The active cell is in the field " + ActiveCell.PivotField.Name) }` |

#### **PivotCell.PivotItem**

返回一个 **PivotItem** 对象，它代表指定区域左上角所在的数据透视表项。

**语法**

**express.PivotItem**

*express*   一个代表 **PivotCell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示包含 Sheet1 中活动单元格的数据透视表数据项的名称.*/ function test(){ Worksheets.Item("Sheet1").Activate() alert("The active cell is in the field " + ActiveCell.PivotField.Name) }` |

#### **PivotCell.PivotRowLine**

在行上返回特定 **PivotCell** 对象的 PivotLine。只读 **PivotLine** 类型。

**语法**

**express.PivotRowLine**

*express*   一个代表 **PivotCell** 对象的变量。

**说明**

如果 PivotCell 在行上，则 **PivotRowLine** 返回行的 **PivotLine** 对象。

如果 PivotCell 在列上，则 **PivotRowLine** 返回一个运行时错误。

如果 PivotCell 在数据区域中，则 **PivotRowLine** 返回对应的行的 **PivotLine** 对象。

#### **PivotCell.PivotTable**

返回一个 **PivotTable** 对象，它代表与 PivotCell 相关联的数据透视表。

**语法**

**express.PivotTable**

*express*   一个代表 **PivotCell** 对象的变量。

**示例**

此示例将 Sheet1 上的数据透视表的当前页设置为名为“Canada”的页。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable pvtTable.PivotFields("Country").CurrentPage = "Canada"` |

此示例确定与“Sales”图表相关联的数据透视表，然后将名为“Oregon”的页设置为数据透视表的当前页。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let objPT = ActiveSheet.Charts("Sales").PivotLayout.PivotTable objPT.PivotFields.Item("State").CurrentPageName = "Oregon"` |

#### **PivotCell.Range**

返回一个 **Range** 对象，它代表指定的 PivotCell 所应用的区域。

**语法**

**express.Range**

*express*   一个代表 **PivotCell** 对象的变量。

**示例**

下例将应用于“Crew”工作表的“自动筛选”的地址保存在变量中。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let rAddress = Worksheets.Item("Crew").AutoFilter.Range.Address()` |

此示例滚动工作簿窗口，直至超链接区域位于活动窗口的左上角。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Workbooks.Item(1).Activate() let hr = ActiveSheet.Hyperlinks.Item(1).Range ActiveWindow.ScrollRow = hr.Row ActiveWindow.ScrollColumn = hr.Column` |

#### **PivotCell.RowItems**

返回一个 **PivotItemList** 集合，该集合对应于分类轴上表示选定单元格的项目。

**语法**

**express.RowItems**

*express*   一个代表 **PivotCell** 对象的变量。

**示例**

本示例判断单元格 B5 中的数据项是否在第一个行字段中的“Inventory”项之下，并通知用户。本示例假定数据透视表位于活动工作表上，并且工作表的列 B 包含数据透视表的行项目。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){  // Determine if there is a match between the item and row field. 									     if(Application.Range("B5").PivotCell.RowItems.Item(1) == "Inventory" ) {         alert("Cell B5 is a member of the 'Inventory' row field.")     }									     else {         alert("Cell B5 is not a member of the 'Inventory' row field.")     } }` |

适用环境：web

适用平台：windows/linux