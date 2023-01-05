**ChartObjects**



由指定的图表工作表、对话框工作表或工作表上的所有 **ChartObject** 对象组成的集合。

**说明**

每个 **ChartObject** 对象都代表一个嵌入图表。**ChartObject** 对象充当 **Chart** 对象的容器。**ChartObject** 对象的属性和方法控制工作表上嵌入图表的外观和大小。**ChartObjects** 集合

**示例**

使用 **ChartObjects** 方法返回 **ChartObjects** 集合。以下示例删除名为“Sheet1”的工作表上的所有嵌入图表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("Sheet1").ChartObjects().Delete()` |

不能使用 **ChartObjects** 集合来调用以下属性和方法：

- **Locked** 属性
- **Placement** 属性
- **PrintObject** 属性

与早期版本不同，**ChartObjects** 集合现在可以读取表示高度、宽度、左对齐和顶对齐的属性。

使用 **Add** 方法可创建一个新的空嵌入图表并将它添加到集合中。使用 **ChartWizard** 方法可添加数据并设置新图表的格式。以下示例创建一个新嵌入图表，然后以折线图形式添加单元格 A1:A20 中的数据。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let ch = Application.Worksheets.Item("Sheet1").ChartObjects().Add(100, 30, 400, 250) ch.Chart.ChartWizard(Worksheets.Item("sheet1").Range("a1:a20"), xlLine     , null, null, null, null, null, "New Chart", null, null, null)` |

使用 **ChartObjects**(*index*)（其中 *index* 是嵌入图表的索引号或名称）可以返回单个对象。以下示例设置名为“Sheet1”的工作表上嵌入图表 Chart 1 中的图表区图案。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.      ChartArea.Format.Fill.Pattern = msoPatternLightDownwardDiagonal` |

**方法**

|                                                              | 名称            | 说明                                 |
| ------------------------------------------------------------ | --------------- | ------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**         | 创建新的嵌入式图表。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**        | 将对象复制到剪贴板。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CopyPicture** | 将所选对象作为图片复制到剪贴板。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**         | 将对象剪切到剪贴板。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**      | 删除对象。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Duplicate**   | 复制对象，并返回对新复制对象的引用。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**        | 从集合中返回一个对象。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**      | 选择对象。                           |

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**        | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**              | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**            | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**             | 返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**               | 返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**             | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectChartObject** | 如果嵌入式图表框不能通过用户界面移动、调整大小或删除，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShapeRange**         | 返回一个 **ShapeRange** 对象，它代表指定的一个或多个对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                | 返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**            | 返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**              | 返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。 |

**成员方法**

#### **ChartObjects.Add**

创建新的嵌入式图表。

**语法**

**express.Add(Left, Width)**

*express*   一个代表 **ChartObjects** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Left*   | 必选          | **Double**   | 以磅为单位给出新对象的初始坐标，该坐标是相对于工作表上单元格 A1 的左上角或图表的左上角的坐标。 |
| *Width*  | 必选          | **Double**   | 以磅为单位给出新对象的初始大小。                             |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例创建新的嵌入式图表。*/ function test() {     let co = Application.Sheets.Item("Sheet1").ChartObjects().Add(50, 40, 200, 100)     co.Chart.ChartWizard(Worksheets.Item("Sheet1").Range("A1:B2"),          xlColumn, 6, xlColumns, 1, 0, 1) }` |

#### **ChartObjects.Copy**

将对象复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **ChartObjects** 对象的变量。

#### **ChartObjects.CopyPicture**

将所选对象作为图片复制到剪贴板。

**语法**

**express.CopyPicture(Appearance, Format)**

*express*   一个代表 **ChartObjects** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**            | **说明**                                |
| ------------ | ------------- | ----------------------- | --------------------------------------- |
| *Appearance* | 可选          | **XlPictureAppearance** | 指定图片的复制方式。默认值为 xlScreen。 |
| *Format*     | 可选          | **XlCopyPictureFormat** | 图片的格式。默认值为 xlPicture。        |

#### **ChartObjects.Cut**

将对象剪切到剪贴板。

**语法**

**express.Cut()**

*express*   一个代表 **ChartObjects** 对象的变量。

**说明**

只可剪切嵌入式图表。

#### **ChartObjects.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **ChartObjects** 对象的变量。

#### **ChartObjects.Duplicate**

复制对象，并返回对新复制对象的引用。

**语法**

**express.Duplicate()**

*express*   一个代表 **ChartObjects** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例复制工作表 Sheet1 上嵌入的第一个图表，然后选择新复制的图表。*/ function test() {       let dChart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Duplicate()     dChart.Select() }  ` |

#### **ChartObjects.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **ChartObjects** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**说明**

对象的文本名称就是 **Name** 和 **Value** 属性的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例激活第一张嵌入式图表。*/ Application.Worksheets.Item("Sheet1").ChartObjects.Item(1).Activate()` |

#### **ChartObjects.Select**

选择对象。

**语法**

**express.Select(Replace)**

*express*   一个代表 **ChartObjects** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | 如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

**成员属性**

#### **ChartObjects.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ChartObjects** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **ChartObjects.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **ChartObjects** 对象的变量。

#### **ChartObjects.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ChartObjects** 对象的变量。

#### **ChartObjects.Height**

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **ChartObjects** 对象的变量。

#### **ChartObjects.Left**

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **ChartObjects** 对象的变量。

#### **ChartObjects.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ChartObjects** 对象的变量。

#### **ChartObjects.ProtectChartObject**

如果嵌入式图表框不能通过用户界面移动、调整大小或删除，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.ProtectChartObject**

*express*   一个代表 **ChartObjects** 对象的变量。

**说明**

将此属性设置为 **True**，嵌入式图表框还是可以通过对象模型进行修改。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).ChartObjects(1).ProtectChartObject = true` |

#### **ChartObjects.ShapeRange**

返回一个 **ShapeRange** 对象，它代表指定的一个或多个对象。只读。

**语法**

**express.ShapeRange**

*express*   一个代表 **ChartObjects** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例创建一个形状区域，该区域代表第一张工作表上的嵌入图表。*/ let sr = Application.Worksheets.Item(1).ChartObjects.ShapeRange` |

#### **ChartObjects.Top**

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **ChartObjects** 对象的变量。

#### **ChartObjects.Visible**

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **ChartObjects** 对象的变量。

#### **ChartObjects.Width**

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **ChartObjects** 对象的变量。

适用环境：web

适用平台：windows/linux