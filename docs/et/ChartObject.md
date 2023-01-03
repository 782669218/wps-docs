**ChartObject**



代表工作表上的嵌入图表。

**说明**

**ChartObject** 对象充当 **Chart** 对象的容器。**ChartObject** 对象的属性和方法控制工作表上嵌入图表的外观和大小。**ChartObject** 对象是 **ChartObjects** 集合的成员。**ChartObjects** 集合包含单一工作表上的所有嵌入图表。

使用 **ChartObjects**(*index*)（其中 *index* 是嵌入图表的索引号或名称）可以返回单个 **ChartObject** 对象。

**示例**
以下示例设置名为“Sheet1”的工作表上嵌入图表 Chart 1 中的图表区图案。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format.Fill.Pattern = msoPatternLightDownwardDiagonal ` |

当选定嵌入图表时，其名称显示在“名称”框中。使用 **Name** 属性可设置或返回 **ChartObject** 对象的名称。以下示例对工作表“Sheet1”上的嵌入图表“Chart 1”使用了圆角。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("Sheet1").ChartObjects("Chart 1").RoundedCorners = true` |

**方法**

|                                                              | 名称             | 说明                                 |
| ------------------------------------------------------------ | ---------------- | ------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**     | 使当前图表成为活动图表。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BringToFront** | 将对象放到 z-次序前面。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**         | 将对象复制到剪贴板。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CopyPicture**  | 将所选对象作为图片复制到剪贴板。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**          | 将对象剪切到剪贴板。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**       | 删除对象。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Duplicate**    | 复制对象，并返回对新复制对象的引用。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**       | 选择对象。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SendToBack**   | 将对象放到 z-次序的后面。            |

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**        | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BottomRightCell**    | 返回一个 **Range** 对象，它代表位于该对象右下角下方的单元格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Chart**              | 返回一个 **Chart** 对象，该对象代表对象中包含的图表。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**            | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**             | 返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**              | 返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**               | 返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Locked**             | 返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**               | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**             | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Placement**          | 返回或设置一个包含 **XlPlacement** 常量的 **Variant** 值，它代表对象附加到它所在的单元格的方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintObject**        | 如果打印文档时也打印指定对象，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectChartObject** | 如果不能通过用户界面对嵌入图表框架执行移动、调整大小或删除操作，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RoundedCorners**     | 如果嵌入式图表使用圆角，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**             | 返回或设置 **Boolean** 值，它确定字体是否是阴影字体或对象是否带有阴影。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShapeRange**         | 返回一个 **ShapeRange** 对象，它代表指定的一个或多个对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                | 返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TopLeftCell**        | 返回一个 **Range** 对象，它代表位于指定对象左上角下方的单元格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**            | 返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**              | 返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ZOrder**             | 返回指定对象的 z-次序位置。**Long** 型，只读。               |

**成员方法**

#### **ChartObject.Activate**

使当前图表成为活动图表。

**语法**

**express.Activate()**

*express*   一个代表 **ChartObject** 对象的变量。

#### **ChartObject.BringToFront**

将对象放到 z-次序前面。

**语法**

**express.BringToFront()**

*express*   一个代表 **ChartObject** 对象的变量。

#### **ChartObject.Copy**

将对象复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **ChartObject** 对象的变量。

#### **ChartObject.CopyPicture**

将所选对象作为图片复制到剪贴板。

**语法**

**express.CopyPicture(Appearance, Format)**

*express*   一个代表 **ChartObject** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**            | **说明**                                |
| ------------ | ------------- | ----------------------- | --------------------------------------- |
| *Appearance* | 可选          | **XlPictureAppearance** | 指定图片的复制方式。默认值为 xlScreen。 |
| *Format*     | 可选          | **XlCopyPictureFormat** | 图片的格式。默认值为 xlPicture。        |

#### **ChartObject.Cut**

将对象剪切到剪贴板。

**语法**

**express.Cut()**

*express*   一个代表 **ChartObject** 对象的变量。

**说明**

只可剪切嵌入式图表。

#### **ChartObject.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **ChartObject** 对象的变量。

#### **ChartObject.Duplicate**

复制对象，并返回对新复制对象的引用。

**语法**

**express.Duplicate()**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例复制工作表 Sheet1 上嵌入的第一个图表，然后选择新复制的图表。*/ function test() {     let dChart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Duplicate()     dChart.Select() }` |

#### **ChartObject.Select**

选择对象。

**语法**

**express.Select(Replace)**

*express*   一个代表 **ChartObject** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | 如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

#### **ChartObject.SendToBack**

将对象放到 z-次序的后面。

**语法**

**express.SendToBack()**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将工作表 Sheet1 上嵌入的第一个图表放到 z-次序的后面。*/ Application.Worksheets.Item("Sheet1").ChartObjects(1).SendToBack()` |

**成员属性**

#### **ChartObject.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **ChartObject.BottomRightCell**

返回一个 **Range** 对象，它代表位于该对象右下角下方的单元格。只读。

**语法**

**express.BottomRightCell**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示 Sheet1 中嵌入的第一个图表右下角单元格的地址。*/ alert("The bottom right corner is over cell " +     Application.Worksheets.Item("Sheet1").ChartObjects(1).BottomRightCell.Address())` |

#### **ChartObject.Chart**

返回一个 **Chart** 对象，该对象代表对象中包含的图表。只读。

**语法**

**express.Chart**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Sheet1 上嵌入的第一个图表添加标题。*/ function test() {     let myChart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart     myChart.HasTitle = true     myChart.ChartTitle.Text = "1995 Rainfall Totals by Month" }` |

#### **ChartObject.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ChartObject** 对象的变量。

#### **ChartObject.Height**

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **ChartObject** 对象的变量。

#### **ChartObject.Index**

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **ChartObject** 对象的变量。

#### **ChartObject.Left**

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例使嵌入图表的左边界与 B 列的左边界对齐。*/ function test() {     let mySheet = Application.Worksheets.Item("Sheet1")     mySheet.ChartObjects(1).Left = mySheet.Columns.Item("B").Left }` |

#### **ChartObject.Locked**

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

*express*   一个代表 **ChartObject** 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**。

#### **ChartObject.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **ChartObject** 对象的变量。

**说明**

此属性对图表对象（嵌入式图表）而言是只读的。

#### **ChartObject.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ChartObject** 对象的变量。

#### **ChartObject.Placement**

返回或设置一个包含 **XlPlacement** 常量的 **Variant** 值，它代表对象附加到它所在的单元格的方式。

**语法**

**express.Placement**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置工作表 Sheet1 上嵌入的第一个图表为可自由浮动（既不随下方单元格移动，也不随其改变大小）。*/ Application.Worksheets.Item("Sheet1").ChartObjects(1).Placement = xlFreeFloating` |

#### **ChartObject.PrintObject**

如果打印文档时也打印指定对象，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintObject**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置当打印工作表 sheet1 时连同嵌入的第一个图表一起打印。*/ Application.Worksheets.Item("Sheet1").ChartObjects(1).PrintObject = true` |

#### **ChartObject.ProtectChartObject**

如果不能通过用户界面对嵌入图表框架执行移动、调整大小或删除操作，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ProtectChartObject**

*express*   一个代表 **ChartObject** 对象的变量。

**说明**

将该属性设置为 **True** 后将无法防止通过对象模型修改嵌入图表框架。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例保护第一张工作表上嵌入的第一个图表。*/ Application.Worksheets.Item(1).ChartObjects(1).ProtectChartObject = true` |

#### **ChartObject.RoundedCorners**

如果嵌入式图表使用圆角，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.RoundedCorners**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例为 Sheet1 的第一个嵌入式图表添加圆角。*/ Application.Worksheets.Item("Sheet1").ChartObjects(1).RoundedCorners = true` |

#### **ChartObject.Shadow**

返回或设置 **Boolean** 值，它确定字体是否是阴影字体或对象是否带有阴影。

**语法**

**express.Shadow**

*express*   一个代表 **ChartObject** 对象的变量。

**说明**

对于 **Font** 对象，该属性在 Microsoft Windows 中无效，但保留其值（可被设置或返回）。

#### **ChartObject.ShapeRange**

返回一个 **ShapeRange** 对象，它代表指定的一个或多个对象。只读。

**语法**

**express.ShapeRange**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例创建一个形状区域，该区域代表第一张工作表上的嵌入图表。*/ let sr = Application.Worksheets.Item(1).ChartObjects.ShapeRange` |

#### **ChartObject.Top**

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **ChartObject** 对象的变量。

#### **ChartObject.TopLeftCell**

返回一个 **Range** 对象，它代表位于指定对象左上角下方的单元格。只读。

**语法**

**express.TopLeftCell**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示 Sheet1 上嵌入的第一张图表左上角下面的单元格的地址。*/ function test() {     alert("The top left corner is over cell " +     Application.Worksheets.Item("Sheet1").ChartObjects(1).TopLeftCell.Address()) }` |

#### **ChartObject.Visible**

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **ChartObject** 对象的变量。

#### **ChartObject.Width**

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **ChartObject** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置嵌入图表的宽度。*/ Application.Worksheets.Item("Sheet1").ChartObjects(1).Width = 360` |

#### **ChartObject.ZOrder**

返回指定对象的 z-次序位置。**Long** 型，只读。

**语法**

**express.ZOrder**

*express*   一个代表 **ChartObject** 对象的变量。

**说明**

在任何对象集合中，z-次序尾端的对象为 *collection*(1)，z-次序前端的对象为 *collection*(*collection*.**Count**)。例如，如果活动工作表中有嵌入图表，z-次序尾端的图表为 `ActiveSheet.ChartObjects(1)`，z-次序前端的图表为 `ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count)`。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示 Sheet1 上嵌入的第一张图表的 z-次序位置。*/ function test() {     alert("The chart's z-order position is " +     Application.Worksheets.Item("Sheet1").ChartObjects(1).ZOrder) }` |

适用环境：web

适用平台：windows/linux