**ListObject**



代表工作表中的表格。

**说明**

**ListObject** 对象是 **ListObjects** 集合的成员。**ListObjects** 集合包含工作表上所有的列表对象。

使用 **Worksheet** 对象的 ListObjects 属性可返回一个 **ListObjects** 集合。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例给活动工作簿的第一张工作表的默认 ListObject 对象添加新的 ListRow 对象。*/ function test(){ let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1") Set oListCol = wrksht.ListObjects.Item(1).ListRows.Add() }` |

**方法**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**        | 删除 **ListObject** 对象并从工作表中清除单元格数据。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ExportToVisio** | 将 **ListObject** 对象导出到 Visio。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Publish**       | 将 **ListObject** 对象发布到运行 Microsoft SharePoint Foundation 的服务器。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Refresh**       | 从运行 Microsoft SharePoint Foundation 的服务器上检索列表的当前数据和架构。此方法只能被用于链接到 SharePoint 网站的列表。如果 SharePoint 网站不可用，调用此方法将返回错误。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Resize**        | **Resize** 方法允许 **ListObject** 对象在新区域内调整大小。不插入或移动单元格。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Unlink**        | 从列表中删除与 Microsoft SharePoint Foundation 网站的链接。返回 **Nothing**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Unlist**        | 从 **ListObject** 对象删除列表功能。使用此方法后，组成列表的单元格区域将变为常规的数据区域。 |

**属性**

|                                                              | 名称                            | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Active**                      | 返回一个 **Boolean** 类型的值，该值表示工作表中的 **ListObject** 对象是否是活动的，即活动的单元格是否在 **ListObject** 对象的区域之内。只读 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlternativeText**             | 返回或设置指定表的描述性（可选）文本字符串。可读/写          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                 | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFilter**                  | 使用“自动筛选”筛选一个列表。只读。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Comment**                     | 返回或设置与列表对象关联的批注。可读/写 **String** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataBodyRange**               | 返回 **Range** 对象，它代表表中值的范围，但不包括标题行。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayName**                 | 返回或设置指定的 **ListObject** 对象的显示名称。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayRightToLeft**          | 如果指定的 **ListObject** 从右到左而不是从左到右显示，则此属性为 **True**。如果对象从左到右显示，则此属性为 **False**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeaderRowRange**              | 返回一个 **Range** 对象，该对象表示列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。）标题行的区域。**Range** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsertRowRange**              | 返回一个 **Range** 对象，该对象代表指定的 **ListObject** 对象的“插入”行（如果有）。**Range** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListColumns**                 | 返回一个 **ListColumns** 集合，该集合代表 **ListObject** 对象中的所有列。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListRows**                    | 返回一个 **ListRows** 对象，该对象代表 **ListObject** 对象中的所有数据行。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                        | 返回或设置一个 **String** 值，它代表 **ListObject** 对象的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **QueryTable**                  | 返回 **QueryTable** 对象，它为 **ListObject** 对象提供了一个列表服务器的链接。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**                       | 返回一个 **Range** 对象，它代表在以上列表中的指定列表对象所应用的区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SharePointURL**               | 返回一个 **String** 类型的数值，该数值表示给定 **ListObject** 对象的 SharePoint 列表的 URL。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAutoFilter**              | 返回一个 **Boolean** 类型的数值，该数值指示是否显示“自动筛选”。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowHeaders**                 | 返回或设置是否显示指定的 **ListObject** 对象的标题信息。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowTableStyleColumnStripes** | 返回或设置指定的 **ListObject** 对象是否使用“列条纹”表样式。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowTableStyleFirstColumn**   | 返回或设置是否显示指定的 **ListObject** 对象的第一列。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowTableStyleLastColumn**    | 返回或设置是否显示指定的 **ListObject** 对象的最后一列。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowTableStyleRowStripes**    | 返回或设置指定的 **ListObject** 对象是否使用“行条纹”表样式。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowTotals**                  | 获取或设置一个 **Boolean** 类型的数值以指示“总计”行是否可见。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sort**                        | 获取或设置 **ListObject** 集合的排序列和排序次序。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceType**                  | 返回一个 **XlListObjectSourceType** 值，它代表列表的当前源。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Summary**                     | 如果指定区域为分级显示的汇总行或汇总列，则该值为 **True**。该区域应为一行或一列。**Variant** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TableStyle**                  | 获取或设置指定的 **ListObject** 对象的表样式。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TotalsRowRange**              | 返回一个 **Range**，代表指定的 **ListObject** 对象的“汇总”行（如果有）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XmlMap**                      | 返回一个 **XmlMap** 对象，该对象表示用于指定表格的架构映射。只读。 |

**成员方法**

#### **ListObject.Delete**

删除 **ListObject** 对象并从工作表中清除单元格数据。

**语法**

**express.Delete()**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

如果列表链接到 SharePoint 网站，删除列表不会影响运行 SharePoint Foundation 的服务器上的数据。对本地列表所做的任何未提交的更改不会发送到 SharePoint 列表（系统不会警告这些未提交的更改将会丢失）。

#### **ListObject.ExportToVisio**

将 **ListObject** 对象导出到 Visio。

**语法**

**express.ExportToVisio()**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.Publish**

将 **ListObject** 对象发布到运行 Microsoft SharePoint Foundation 的服务器。

**语法**

**express.Publish(Target, LinkSource)**

*express*   一个代表 **ListObject** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                   |
| ------------ | ------------- | ------------ | ------------------------------------------ |
| *Target*     | 必选          | **Variant**  | 包含一个 String 值数组，如“备注”部分所述。 |
| *LinkSource* | 必选          | **Boolean**  |                                            |

**返回值**

一个字符串，代表 SharePoint 网站上已发布列表的 URL。

**说明**

*Target* 参数包含一个 **String** 元素数组，如下表中所述：

| 元素号 | 内容                    |
| ------ | ----------------------- |
| 0      | SharePoint 服务器的 URL |
| 1      | ListName（显示名称）    |
| 2      | 列表的说明。可选。      |

如果 **ListObject** 对象目前没有链接到 SharePoint 网站上的列表，在将 *LinkSource* 设置为 **True** 时，将在指定的 SharePoint 网站上创建新的列表。如果 **ListObject** 对象目前已经链接到 SharePoint 网站上的列表，在将 *LinkSource* 参数设置为 **True** 时，将替换现有的链接（只能将列表链接到一个 SharePoint 网站）。如果 **ListObject** 对象目前没有链接，在将 *LinkSource* 设置为 **False** 时，将保留 **ListObject** 对象不进行链接。如果 **ListObject** 对象目前已经链接到 SharePoint 网站，在将 *LinkSource* 设置为 **False** 时，将保持 **ListObject** 对象与当前 SharePoint 网站的链接。

#### **ListObject.Refresh**

从运行 Microsoft SharePoint Foundation 的服务器上检索列表的当前数据和架构。此方法只能被用于链接到 SharePoint 网站的列表。如果 SharePoint 网站不可用，调用此方法将返回错误。

**语法**

**express.Refresh()**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

对 **Refresh** 方法的调用不会将更改提交到 ET 工作簿内的列表中。在 ET 中，当调用 **Refresh** 方法时，将放弃列表中未提交的更改。

#### **ListObject.Resize**

**Resize** 方法允许 **ListObject** 对象在新区域内调整大小。不插入或移动单元格。

**语法**

**express.Resize(Range)**

*express*   一个代表 **ListObject** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明** |
| -------- | ------------- | ------------ | -------- |
| *Range*  | 必选          | **Range**    | 新区域。 |

**说明**

对于链接到正在运行 Microsoft SharePoint Foundation 的服务器的表格，可使用此方法调整列表大小，方法是提供一个仅在所包含的行数方面与 **ListObject** 的当前区域不同的 *Range* 参数。如果通过添加或删除列（在 *Range* 参数中）来调整链接到 SharePoint Foundation 的列表的大小，将导致运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例使用了 Resize 方法来调整活动工作簿的 Sheet1 中默认 ListObject 对象的大小。*/ function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let objListObj = wrksht.ListObjects.Item(1)                 objListObj.Resize (Range("A1:B10"))  }` |

#### **ListObject.Unlink**

从列表中删除与 Microsoft SharePoint Foundation 网站的链接。返回 **Nothing**。

**语法**

**express.Unlink()**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

调用此方法后，列表被取消链接，无法再恢复。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例从 SharePoint 网站上删除了列表链接。*/ function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let objListObj = wrksht.ListObjects.Item(1)    objListObj.Unlink() }` |

#### **ListObject.Unlist**

从 **ListObject** 对象删除列表功能。使用此方法后，组成列表的单元格区域将变为常规的数据区域。

**语法**

**express.Unlist()**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

运行此方法可保留工作表中的单元格数据、格式和公式。**“汇总行”**也保留不变。此方法可删除 Microsoft SharePoint Foundation 网站的所有链接。**“自动筛选”**也从列表中删除。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例将列表功能从工作表上的某个列表中删除。*/ function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let objListObj = wrksht.ListObjects.Item(1)     objListObj.Unlist() }` |

**成员属性**

#### **ListObject.Active**

返回一个 **Boolean** 类型的值，该值表示工作表中的 **ListObject** 对象是否是活动的，即活动的单元格是否在 **ListObject** 对象的区域之内。只读 **Boolean** 类型。

**语法**

**express.Active**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

因为 **ListObject** 对象没有 **Activate** 方法，所以只能通过激活列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。）中某个单元格区域的方式激活 **ListObject** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例激活了活动工作簿第一个工作表的默认 ListObject 对象的列表。调用 Range 对象的 Activate 方法时不提供单元格参数，将激活列表的整个区域。*/ function test(){ let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let objList = wrksht.ListObjects.Item(1)     objList.Range.Activate()      let MakeListActive = objList.Active }` |

#### **ListObject.AlternativeText**

返回或设置指定表的描述性（可选）文本字符串。可读/写

**语法**

**express.AlternativeText**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

**AlternativeText** 属性的值对应于**“可选文字”**对话框（右键单击表、指向**“表”**，然后单击**“可选文字”**即可显示该对话框）中**“标题”**框的设置。

#### **ListObject.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ListObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **ListObject.AutoFilter**

使用“自动筛选”筛选一个列表。只读。

**语法**

**express.AutoFilter**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.Comment**

返回或设置与列表对象关联的批注。可读/写 **String** 类型。

**语法**

**express.Comment**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

批注文本不能超过 255 个字符。

#### **ListObject.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ListObject.DataBodyRange**

返回 **Range** 对象，它代表表中值的范围，但不包括标题行。只读。

**语法**

**express.DataBodyRange**

*express*   一个代表 **ListObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例选定列表中的活动数据区域。*/ function test(){ Application.Worksheets.Item("Sheet1").Activate() ActiveWorksheet.ListObjects.Item(1).DataBodyRange.Select() }` |

#### **ListObject.DisplayName**

返回或设置指定的 **ListObject** 对象的显示名称。可读/写 **String** 类型。

**语法**

**express.DisplayName**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.DisplayRightToLeft**

如果指定的 **ListObject** 从右到左而不是从左到右显示，则此属性为 **True**。如果对象从左到右显示，则此属性为 **False**。**Boolean** 类型，只读。

**语法**

**express.DisplayRightToLeft**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.HeaderRowRange**

返回一个 **Range** 对象，该对象表示列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。）标题行的区域。**Range** 类型，只读。

**语法**

**express.HeaderRowRange**

*express*   一个代表 **ListObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例激活了由活动工作簿第一个工作表中默认 ListObject 对象的 HeaderRowRange 属性指定的区域。*/ function test(){ let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let objList = wrksht.ListObjects.Item(1)     let objListRng = objList.HeaderRowRange     objListRng.Activate() }` |

#### **ListObject.InsertRowRange**

返回一个 **Range** 对象，该对象代表指定的 **ListObject** 对象的“插入”行（如果有）。**Range** 类型，只读。

**语法**

**express.InsertRowRange**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

如果由于列表?（列表：包含相关数据的一系列行，或使用 “创建列表” 命令作为数据表指定给函数的一系列行。）处于非活动状态而使得“插入”行不可见，则返回 **Nothing** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例激活了由活动工作簿第一个工作表内默认 ListObject 对象中 InsertRowRange 指定的区域。*/ function test(){   let wrksht = Application.ActiveWorkbook.Worksheets.Item(1)     let objList = wrksht.ListObjects.Item(1)     let objListRng = objList.InsertRowRange     if(objListRng){         ActivateInsertRow = false 	}     else{         objListRng.Activate()         ActivateInsertRow = true     } }` |

#### **ListObject.ListColumns**

返回一个 **ListColumns** 集合，该集合代表 **ListObject** 对象中的所有列。只读。

**语法**

**express.ListColumns**

*express*   一个代表 **ListObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例显示了通过对 ListColumns 属性的调用所创建的 ListColumns 集合对象中第二列的名称。要让本代码运行，Sheet1 工作表必须包含一个表格。*/ function test(){ let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let objListObj = wrksht.ListObjects.Item(1)     let objListCols = objListObj.ListColumns     alert(objListCols.Item(2).Name) }` |

#### **ListObject.ListRows**

返回一个 **ListRows** 对象，该对象代表 **ListObject** 对象中的所有数据行。只读。

**语法**

**express.ListRows**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

返回的 **ListRows** 对象不包括标题行、总计行或插入行。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例删除了由 ListRows 集合中的数字指定的行，该集合是通过调用 ListRows 属性创建的。*/ function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let objListObj = wrksht.ListObjects.Item(1)     let objListRows = objListObj.ListRows   	if((iRowNumber != 0) && (iRowNumber < objListRows.Count - 1)){         objListRows.Item(iRowNumber).Delete()     } }` |

#### **ListObject.Name**

返回或设置一个 **String** 值，它代表 **ListObject** 对象的名称。

**语法**

**express.Name**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

此名称只用作 **ListObjects** 集合对象的 **Item** 属性的唯一标识符。只能通过对象模型设置此属性。

默认情况下，每个 **ListObject** 对象名称都以“List”开头，后面接着数字（无空格）。如果试图将 **Name** 属性设置为其他 **ListObject** 对象已经使用的名称，将产生运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动工作簿的 sheet1 中的默认 ListObject 对象的名称。*/ function test(){   let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let oListObj = wrksht.ListObjects.Item(1)    alert (oListObj.Name)   }` |

#### **ListObject.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.QueryTable**

返回 **QueryTable** 对象，它为 **ListObject** 对象提供了一个列表服务器的链接。只读。

**语法**

**express.QueryTable**

*express*   一个代表 **ListObject** 对象的变量。

**示例**

下例创建一个到 SharePoint 网站的链接，并将名为 `List1` 的 **ListObject** 对象发布到服务器。另外，还创建了对列表对象的 **QueryTable** 对象的引用，并将 **QueryTable** 对象的 **MaintainConnection** 属性设为 **True**，以便保持服务器与 SharePoint 网站之间的连接。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1") let objListObj = wrksht.ListObjects.Item(1)  let arr = [] arr[0] = "0" arr[1] = "http://myteam/project1" arr[2] = "1" arr[3] = "List1"     let strSTSConnection = objListObj.Publish(arr, true)     let objQryTbl = objListObj.QueryTable     objQryTbl.MaintainConnection = true }` |

#### **ListObject.Range**

返回一个 **Range** 对象，它代表在以上列表中的指定列表对象所应用的区域。

**语法**

**express.Range**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.SharePointURL**

返回一个 **String** 类型的数值，该数值表示给定 **ListObject** 对象的 SharePoint 列表的 URL。**String** 类型，只读。

**语法**

**express.SharePointURL**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

如果列表未与 SharePoint 网站链接，则访问该属性会产生运行时错误。

**示例**

下例为 **Publish** 方法的 *Target* 参数设置了元素，以便将 **ListObject** 对象推送到 SharePoint 网站。此代码示例使用 **SharePointURL** 属性将 URL 分配给数组，使用 **Name** 属性分配列表的名称。然后使用 **Publish** 方法将数组中的信息传递到 SharePoint 网站。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let objListObj = wrksht.ListObjects.Item(1)      let arr = []     arr[0] = "0"     arr[1] = objListObj.SharePointURL     arr[2] = "1"     arr[3] = objListObj.Name         let strSTSConnection = objListObj.Publish(arr, true) }` |

#### **ListObject.ShowAutoFilter**

返回一个 **Boolean** 类型的数值，该数值指示是否显示“自动筛选”。**Boolean** 类型，可读写。

**语法**

**express.ShowAutoFilter**

*express*   一个代表 **ListObject** 对象的变量。

**说明**

对于新的 **ListObject** 对象，**ShowAutoFilter** 属性默认为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例显示了对活动工作簿的 Sheet 1 中默认列表进行的 ShowAutoFilter 属性设置。*/ function test(){ let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1") let oListCol = wrksht.ListObjects.Item(1) Debug.Print (oListCol.ShowAutoFilter) }` |

#### **ListObject.ShowHeaders**

返回或设置是否显示指定的 **ListObject** 对象的标题信息。可读/写 **Boolean** 类型。

**语法**

**express.ShowHeaders**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.ShowTableStyleColumnStripes**

返回或设置指定的 **ListObject** 对象是否使用“列条纹”表样式。可读/写 **Boolean** 类型。

**语法**

**express.ShowTableStyleColumnStripes**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.ShowTableStyleFirstColumn**

返回或设置是否显示指定的 **ListObject** 对象的第一列。可读/写 **Boolean** 类型。

**语法**

**express.ShowTableStyleFirstColumn**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.ShowTableStyleLastColumn**

返回或设置是否显示指定的 **ListObject** 对象的最后一列。可读/写 **Boolean** 类型。

**语法**

**express.ShowTableStyleLastColumn**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.ShowTableStyleRowStripes**

返回或设置指定的 **ListObject** 对象是否使用“行条纹”表样式。可读/写 **Boolean** 类型。

**语法**

**express.ShowTableStyleRowStripes**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.ShowTotals**

获取或设置一个 **Boolean** 类型的数值以指示“总计”行是否可见。**Boolean** 类型，可读写。

**语法**

**express.ShowTotals**

*express*   一个代表 **ListObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示了对活动工作簿的 Sheet1 中默认列表的 ShowTotals 属性的当前设置。*/ function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let oListObj = wrksht.ListObjects.Item(1)    Debug.Print (oListObj.ShowTotals) }` |

#### **ListObject.Sort**

获取或设置 **ListObject** 集合的排序列和排序次序。

**语法**

**express.Sort**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.SourceType**

返回一个 **XlListObjectSourceType** 值，它代表列表的当前源。

**语法**

**express.SourceType**

*express*   一个代表 **ListObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例代码返回 XlListObjectSourceType 常量，该常量表示活动工作簿的 Sheet1 上的默认列表的源。*/ function test(){   let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")    let oListObj = wrksht.ListObjects.Item(1)        Debug.Print (oListObj.SourceType) }` |

#### **ListObject.Summary**

如果指定区域为分级显示的汇总行或汇总列，则该值为 **True**。该区域应为一行或一列。**Variant** 类型，只读。

**语法**

**express.Summary**

*express*   一个代表 **ListObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/**/ function test(){ let rows = Application.Worksheets.Item("Sheet1").Rows.Item(4) if(rows.Summary == true){     rows.Font.Bold = true     rows.Font.Italic = true } }` |

#### **ListObject.TableStyle**

获取或设置指定的 **ListObject** 对象的表样式。可读/写 **Variant** 类型。

**语法**

**express.TableStyle**

*express*   一个代表 **ListObject** 对象的变量。

#### **ListObject.TotalsRowRange**

返回一个 **Range**，代表指定的 **ListObject** 对象的“汇总”行（如果有）。只读。

**语法**

**express.TotalsRowRange**

*express*   一个代表 **ListObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例代码返回了活动工作簿的 Sheet1 中默认列表的“汇总”行的地址。该代码可在“汇总”行尚未显示的情况下显示该行。*/ function test(){ let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let objListObj = wrksht.ListObjects.Item(1)     objListObj.ShowTotals = true     alert (objListObj.TotalsRowRange.Address()) }` |

#### **ListObject.XmlMap**

返回一个 **XmlMap** 对象，该对象表示用于指定表格的架构映射。只读。

**语法**

**express.XmlMap**

*express*   一个代表 **ListObject** 对象的变量。

适用环境：web

适用平台：windows/linux