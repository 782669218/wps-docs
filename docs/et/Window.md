**Window**



代表窗口。

**说明**

许多工作表特征（如滚动条和标尺）实际上是窗口的属性。**Window** 对象是 **Windows** 集合的成员。**Application** 对象的 **Windows** 集合包含应用程序中的所有窗口，而 **Workbook** 对象的 **Windows** 集合只包含指定工作簿中的窗口。

## 示例

使用 **Windows**(*index*)（其中 *index* 是窗口名称或索引号）可返回一个 **Window** 对象。下例最大化活动窗口。

| 示例代码复制                                            |
| ------------------------------------------------------- |
| `Application.Windows.Item(1).WindowState = xlMaximized` |

注意，活动窗口总是 `Windows(1)`。

窗口题主是窗口未处于最大化时窗口最上方标题栏中显示的文本。该题注还会显示在**“窗口”**菜单最下方的打开文件列表中。使用 **Caption** 属性可设置或返回窗口题注。更改窗口题注不会更改工作簿的名称。下例为“Book1.xls:1”窗口中显示的工作表关闭单元格网格线。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Windows.Item("book1.xls":1).DisplayGridlines = false` |

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**         | 将窗口提到 z-次序的最前面。返回值类型为Variant。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ActivateNext**     | 激活指定窗口，并将其移到窗口 z-次序的末尾。返回值类型为Variant。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ActivatePrevious** | 激活指定窗口，并将其移到窗口 z-次序的末尾。返回值类型为Variant。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Close**            | 关闭对象。返回值：如果该方法成功关闭对象，则为 **True**，否则为 **False**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NewWindow**        | 新建一个窗口或者创建指定窗口的副本。返回值类型为Window。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RangeFromPoint**   | 新建一个窗口或者创建指定窗口的副本。返回值类型为Object。     |

**属性**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveCell**     | 返回一个 **Range** 对象，它代表活动窗口（最上方的窗口）或指定窗口中的活动单元格。如果窗口中没有显示工作表，此属性无效。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveSheet**    | 返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 **Nothing**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**    | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**        | 返回或设置一个 **Variant** 值，它代表文档窗口标题栏中显示的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**         | 返回或设置一个 Double 值，它代表窗口的高度（以磅为单位）。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**          | 返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**           | 返回一个 **Double** 值，它代表从客户区左边缘到窗口左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RangeSelection** | 返回一个 **Range** 对象，该对象表示指定窗口中工作表上的选定单元格，即使工作表上一个图形对象是活动或选定的。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Selection**      | 对于 **Windows** 对象，返回一个指定的窗口。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SheetViews**     | 返回指定窗口的 **SheetViews** 对象。只读。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**            | 返回或设置一个 **Double** 值，它代表从窗口上边缘到使用区域（在菜单、任何停放在顶端的工具栏和编辑栏下方）上边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**           | 返回或设置一个代表窗口类型的 **XlWindowType** 值。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UsableHeight**   | 返回在应用程序窗口区域中一个窗口能占有的最大高度（以磅为单位）。**Double** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UsableWidth**    | 返回在应用程序窗口区域中一个窗口能占有的最大宽度（以磅为单位）。**Double** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**        | 返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VisibleRange**   | 返回一个 **Range** 对象，它代表显示在窗口或窗格中的单元格区域。如果列或行只显示了一部分，则说明它是包括在区域之内的。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**          | 返回或设置一个 **Double** 值，它代表窗口的宽度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WindowState**    | 返回或设置窗口的状态。**XlWindowState** 类型，可读写。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Zoom**           | 返回或设置一个 **Variant** 值，它代表窗口的显示大小，以百分数表示（100 表示正常大小，200 表示双倍大小，以此类推）。 |

**成员方法**

#### **Window.Activate**

将窗口提到 z-次序的最前面。返回值类型为Variant。

**语法**

**express.Activate()**

*express*   一个代表 **Window** 对象的变量。

**说明**

这样不会运行可能附加在工作簿上的 Auto_Activate 或 Auto_Deactivate 宏（使用 **RunAutoMacros** 方法可运行这些宏）。

#### **Window.ActivateNext**

激活指定窗口，并将其移到窗口 z-次序的末尾。返回值类型为Variant。

**语法**

**express.ActivateNext()**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口移到窗口 z-次序的末尾。*/ Application.ActiveWindow.ActivateNext()` |

#### **Window.ActivatePrevious**

激活指定窗口，并将其移到窗口 z-次序的末尾。返回值类型为Variant。

**语法**

**express.ActivatePrevious()**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例激活窗口 z-次序末尾的窗口。*/ Application.ActiveWindow.ActivatePrevious()` |

#### **Window.Close**

关闭对象。返回值：如果该方法成功关闭对象，则为 **True**，否则为 **False**。

**语法**

**express.Close(SaveChanges, Filename, RouteWorkbook)**

*express*   一个代表 **Window** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SaveChanges*   | 可选          | **Variant**  | 如果工作簿中没有改动，则忽略此参数。如果工作簿中有改动但工作簿显示在其他打开的窗口中，则忽略此参数。如果工作簿中有改动且工作簿未显示在任何其他打开的窗口中，则由此参数指定是否应保存更改。如果设为 True，则保存对工作簿所做的更改。如果工作簿尚未命名，则使用 Filename。如果省略 Filename，则要求用户提供文件名。 |
| *Filename*      | 可选          | **Variant**  | 以此文件名保存所做的更改。                                   |
| *RouteWorkbook* | 可选          | **Variant**  | 如果工作簿不需要传送给下一个收件人（没有传送名单或已经传送），则忽略此参数。否则，ET 根据此参数的值传送工作簿。如果设为 True，则将工作簿传送给下一个收件人。如果设为 False，则不发送工作簿。如果忽略，则要求用户确认是否发送工作簿。 |

**说明**

使用 **RunAutoMacros** 方法可运行自动关闭宏。

#### **Window.NewWindow**

新建一个窗口或者创建指定窗口的副本。返回值类型为Window。

**语法**

**express.NewWindow()**

*express*   一个代表 **Window** 对象的变量。

#### **Window.RangeFromPoint**

新建一个窗口或者创建指定窗口的副本。返回值类型为Object。

**语法**

**express.RangeFromPoint(x, y)**

*express*   一个代表 **Window** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                   |
| -------- | ------------- | ------------ | ---------------------------------------------------------- |
| *x*      | 必选          | **Long**     | 表示从顶部开始到屏幕左边缘的水平距离的值（以像素为单位）。 |
| *y*      | 必选          | **Long**     | 表示从左侧开始到屏幕顶部的垂直距离的值（以像素为单位）。   |

**成员属性**

#### **Window.ActiveCell**

返回一个 **Range** 对象，它代表活动窗口（最上方的窗口）或指定窗口中的活动单元格。如果窗口中没有显示工作表，此属性无效。只读。

**语法**

**express.ActiveCell**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果不指定对象识别符，此属性返回活动窗口中的活动单元格。

请仔细区分活动单元格和选定区域。活动单元格为选定区域内部的一个单元格。而选定区域可以包含多个单元格，但只有一个单元格为活动单元格。

下列表达式都是返回活动单元格，并且都是等效的。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveCell Application.ActiveWindow.ActiveCell` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在消息框中显示活动单元格的值。由于如果活动表不是工作表则 ActiveCell 属性无效，所以此示例使用 ActiveCell 属性之前先激活 Sheet1。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     alert(ActiveCell.Value()) }  /*此示例更改活动单元格的字体格式设置。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     let activeCell = ActiveCell.Font     activeCell.Bold = true     activeCell.Italic = true }` |

#### **Window.ActiveSheet**

返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 **Nothing**。

**语法**

**express.ActiveSheet**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果不指定对象识别符，则此属性返回活动工作簿中的活动工作表。

如果某个工作簿出现在若干个窗口中，那么该工作簿的 **ActiveSheet** 属性在不同窗口中可能不同。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示活动工作表的名称。*/ alert("The name of the active sheet is " + Application.ActiveSheet.Name)` |

#### **Window.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     if (Application.ActiveWorkbook.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **Window.Caption**

返回或设置一个 **Variant** 值，它代表文档窗口标题栏中显示的名称。

**语法**

**express.Caption**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果设置了该名称，就可以使用该名称作为 **Windows** 集合的索引（如示例中的所述）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将活动工作簿中第一个窗口的名称设置为“Consolidated Balance Sheet”。然后该名称将用作 Windows 集合中该窗口的索引。*/ function test() {     Application.ActiveWorkbook.Windows.Item(1).Caption = "Consolidated Balance Sheet"     Application.ActiveWorkbook.Windows.Item("Consolidated Balance Sheet").ActiveSheet.Calculate() }` |

#### **Window.Height**

返回或设置一个 Double 值，它代表窗口的高度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **Window** 对象的变量。

**说明**

使用 **UsableHeight** 属性可确定窗口的最大尺寸。如果窗口已被最大化或最小化，则不能设置此属性。使用 **WindowState** 属性可确定窗口的状态。

#### **Window.Index**

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **Window** 对象的变量。

#### **Window.Left**

返回一个 **Double** 值，它代表从客户区左边缘到窗口左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **Window** 对象的变量。

#### **Window.RangeSelection**

返回一个 **Range** 对象，该对象表示指定窗口中工作表上的选定单元格，即使工作表上一个图形对象是活动或选定的。只读。

**语法**

**express.RangeSelection**

*express*   一个代表 **Window** 对象的变量。

**说明**

当工作表中已选定一个图形对象，**Selection** 属性返回的是一个图形对象而不是一个 **Range** 对象； **RangeSelection** 属性将返回在图形对象被选定之前选定的单元格区域。

当工作表中选定的是一个单元格区域时（不是一个图形对象），该属性和 **Selection** 属性将返回相同的值。

如果指定窗口中的活动表不是一个工作表，则该属性无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示在当前窗口的工作表中选定的单元格区域的地址。*/ alert(Application.ActiveWindow.RangeSelection.Address())` |

#### **Window.Selection**

对于 **Windows** 对象，返回一个指定的窗口。

**语法**

**express.Selection**

*express*   一个代表 **Window** 对象的变量。

**说明**

返回的对象类型取决于当前所选内容（例如，如果选择了单元格，此属性将返回 **Range** 对象）。如果未选择任何内容，**Selection** 属性将返回 **Nothing**。

在不使用对象识别符的情况下，使用此属性等效于使用 `Application.Selection`。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例清空 Sheet1 的选定对象（假定选定对象为单元格区域）。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.Selection.Clear() }  /*本示例显示选定对象的 JSAPI 对象类型。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     alert("The selection object type is "+ typeof(Application.Selection)) }` |

#### **Window.SheetViews**

返回指定窗口的 **SheetViews** 对象。只读。

**语法**

**express.SheetViews**

*express*   一个代表 **Window** 对象的变量。

#### **Window.Top**

返回或设置一个 **Double** 值，它代表从窗口上边缘到使用区域（在菜单、任何停放在顶端的工具栏和编辑栏下方）上边缘的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **Window** 对象的变量。

**说明**

无法对最大化窗口设置此属性。使用 **WindowState** 属性可返回或设置窗口的状态。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例水平排列第一个窗口和第二个窗口，也就是说，每个窗口占用可用的垂直空间的一半，占用应用程序窗口的工作区域的所有水平空间。若要运行该示例，必须只打开两个工作表窗口。*/ function test() {     Application.Windows.Arrange(-4128)     let ah = Application.Windows.Item(1).Height                      // available height     let aw = Application.Windows.Item(1).Width + Application.Windows.Item(2).Width    // available width      let wintop1 = Application.Windows.Item(1)     wintop1.Width = aw     wintop1.Height = ah / 2     wintop1.Left = 0      let wintop2 = Application.Windows.Item(2)     wintop2.Width = aw     wintop2.Height = ah / 2     wintop2.Top = ah / 2     wintop2.Left = 0 }` |

#### **Window.Type**

返回或设置一个代表窗口类型的 **XlWindowType** 值。

**语法**

**express.Type**

*express*   一个代表 **Window** 对象的变量。

#### **Window.UsableHeight**

返回在应用程序窗口区域中一个窗口能占有的最大高度（以磅为单位）。**Double** 型，只读。

**语法**

**express.UsableHeight**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将活动窗口扩展为可用的最大值（假定指定窗口尚未最大化）。*/ function test() {     Application.ActiveWindow.WindowState = xlNormal     Application.ActiveWindow.Top = 1     Application.ActiveWindow.Left = 1     Application.ActiveWindow.Height = Application.UsableHeight     Application.ActiveWindow.Width = Application.UsableWidth }` |

#### **Window.UsableWidth**

返回在应用程序窗口区域中一个窗口能占有的最大宽度（以磅为单位）。**Double** 型，只读。

**语法**

**express.UsableWidth**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将活动窗口扩展为可用的最大值（假定指定窗口尚未最大化）。*/ function test() {     Application.ActiveWindow.WindowState = xlNormal     Application.ActiveWindow.Top = 1     Application.ActiveWindow.Left = 1     Application.ActiveWindow.Height = Application.UsableHeight     Application.ActiveWindow.Width = Application.UsableWidth }` |

#### **Window.Visible**

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **Window** 对象的变量。

#### **Window.VisibleRange**

返回一个 **Range** 对象，它代表显示在窗口或窗格中的单元格区域。如果列或行只显示了一部分，则说明它是包括在区域之内的。只读。

**语法**

**express.VisibleRange**

*express*   一个代表 **Window** 对象的变量。

#### **Window.Width**

返回或设置一个 **Double** 值，它代表窗口的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **Window** 对象的变量。

**说明**

使用 **UsableWidth** 属性可确定窗口的最大尺寸。如果窗口已被最大化或最小化，则不能设置此属性。使用 **WindowState** 属性可确定窗口的状态。

#### **Window.WindowState**

返回或设置窗口的状态。**XlWindowState** 类型，可读写。

**语法**

**express.WindowState**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 ET 应用程序窗口最大化。*/ Application.WindowState = xlMaximized  /*本示例将活动窗口扩展为可用的最大值（假定指定窗口尚未最大化）。*/ function test() {     Application.ActiveWindow.WindowState = xlNormal     Application.ActiveWindow.Top = 1     Application.ActiveWindow.Left = 1     Application.ActiveWindow.Height = Application.UsableHeight     Application.ActiveWindow.Width = Application.UsableWidth }` |

#### **Window.Zoom**

返回或设置一个 **Variant** 值，它代表窗口的显示大小，以百分数表示（100 表示正常大小，200 表示双倍大小，以此类推）。

**语法**

**express.Zoom**

*express*   一个代表 **Window** 对象的变量。

**说明**

将此属性设为 **True**，可将窗口大小设置成与当前选定区域相适应的大小。

本功能仅对窗口中当前的活动工作表起作用。若要对其他工作表使用此属性，必须先激活工作表。

适用环境：web

适用平台：windows/linux