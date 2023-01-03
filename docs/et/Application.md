**Application**



代表整个 ET 应用程序，它是整个应用程序api对象树的根对象。

**说明**

整个et应用程序的api对象结构是一个树状结构，而**Application**对象是树状结构的根对象，同时，**Application**对象也提供了对于应用程序相关的各种访问接口，例如：

- 应用程序的设置选项、环境、版本号等相关信息
- 一些常见的属性，如**ActiveCell**、**ActiveWorkbook**、**ActiveSheet**等

## 示例

使用 **Application** 属性可返回 **Application** 对象。下例对 **Application** 对象应用 **Windows** 属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let app = Application.Application;     alert(app.Windows.Count); }` |

下例在 ET 中打开指定的工作簿。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     let docPath = "d:/abc.xlsx";     Application.Workbooks.Open(docPath); } ` |

**方法**

|                                                              | 名称         | 说明                                                         |
| ------------------------------------------------------------ | ------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Evaluate** | 将一个 ET 名称转换为一个对象或者一个值。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Goto**     | 选定任意工作簿中的任意区域，并且如果该工作簿未处于活动状态，就激活该工作簿。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InputBox** | 显示一个接收用户输入的对话框。返回此对话框中输入的信息。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Quit**     | 退出ET                                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Run**      | 运行一个宏或者调用一个函数。该方法可用于运行宏编辑器编写的宏，或者运行 DLL 或 XLL 中的函数。 |

**属性**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveCell**        | 返回一个 Range 对象，该对象代表活动窗口（顶部窗口）或指定窗口中的活动单元格。如果窗口中没有显示工作表，此属性无效。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveSheet**       | 返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 **null**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveWindow**      | 返回一个 **Window**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Window/Window%20.htm#jsObject_Window)对象，该对象代表活动窗口（顶部窗口）。只读。如果没有打开的窗口，则返回 **null**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveWorkbook**    | 返回一个 **Workbook**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)对象，该对象代表活动窗口（顶部窗口）中的工作簿。只读。如果没有打开的窗口，或者“信息”窗口或“剪贴板”窗口为活动窗口，则返回 **null**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**       | 该属性返回一个代表 ET 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Build**             | 返回 ET 内部版本号。**Number** 类型，只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**           | 返回或设置一个 **String** 值，它代表出现在 ET 主窗口标题栏中显示的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Cells**             | 返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表活动工作表中的所有单元格。如果活动文档不是工作表，则此属性无效。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Columns**           | 返回一个 [Range](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表活动工作表中的所有列。如果活动文档不是工作表，则 Columns 属性无效。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Dialogs**           | 返回一个 **Dialogs**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Dialogs/Dialogs%20.htm#jsObject_Dialogs)集合，该集合代表所有内置对话框。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**            | 返回一个 **Double** 值，它代表主应用程序窗口的高度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hwnd**              | 返回一个 **Number** 类型的值，该值表示 ET 窗口的最高级别的窗口句柄。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**              | 返回或设置一个 **Double** 值，它代表从屏幕左边缘到 ET 主窗口左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**              | 返回一个代表对象名称的 **String** 值。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Names**             | 返回一个 **Names**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Names/Names%20.htm#jsObject_Names)集合，该集合代表活动工作簿中的所有名称。只读 **Names**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Names/Names%20.htm#jsObject_Names)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NewWorkbook**       | 返回一个 **NewFile** 对象。                                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**              | 返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rows**              | 返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表活动工作表中的所有行。如果活动文档不是工作表，则 **Rows** 属性失效。**Range** 对象，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Selection**         | 为 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象返回在活动窗口中选定的对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**               | 返回或设置一个 **Double** 值，它代表从屏幕上边缘到 ET 主窗口上边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**             | 返回一个代表应用程序名称的 **String** 值。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Version**           | 返回一个 **String** 值，它代表 ET 版本号。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**             | 返回或设置一 **个 Double** 值，它代表应用程序窗口的宽度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WindowState**       | 返回或设置窗口的状态。[XlWindowState ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlWindowState%20%E6%9E%9A%E4%B8%BE.html)类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Windows**           | 返回一个 **Windows** 集合，它代表所有工作簿中的所有窗口。**Windows**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Windows/Windows%20.htm#jsObject_Windows)对象，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Workbooks**         | 返回一个 **Workbooks**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbooks/Workbooks%20.htm#jsObject_Workbooks)集合，该集合表示所有打开的工作簿。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WorksheetFunction** | 返回**WorksheetFunction**对象。只读。                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Worksheets**        | 对于 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，返回一个 [Sheets ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)集合，它代表活动工作簿中的所有工作表。对于 **Workbook**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)对象，返回一个 [Sheets ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)集合，它代表指定工作簿中的所有工作表。[Sheets](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)对象，只读。 |

**成员方法**

#### **Application.Evaluate**

将一个 ET 名称转换为一个对象或者一个值。

**语法**

**express.Evaluate(Name)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                           |
| -------- | ------------- | ------------ | ---------------------------------- |
| *Name*   | 必选          | **Variant**  | 公式或使用 ET 命名约定的对象名称。 |

**说明**

该方法可使用下列 ET 名称类型：

- A1 格式引用。可以通过 A1 格式表示法引用单个单元格。所有引用均视为绝对引用。
- 区域。在引用中可以使用区域、交集和联合运算符（分别为冒号、空格和逗号）。
- 定义的名称。可用宏语言指定任何名称。
- 外部引用。可以使用 ! 运算符引用另一工作簿中的单元格或已定义的名称，例如，Evaluate("[BOOK1.XLS]Sheet1!A1")。
- 图表对象。可以指定任何图表对象名称（如“Legend”、“Plot Area”或“Series 1”），以访问该对象的属性和方法。例如，Charts("Chart1").Evaluate("Legend").Font.Name 返回图例中所用字体的名称。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 用字符串参数调用 Evaluate 方法 */ function test() {     Application.Evaluate("A1").Value2 = 25     let trigVariable = Application.Evaluate("SIN(45)")     let firstCellInSheet = Application.Workbooks.Item("工作簿1.xlsx").Sheets.Item(1).Evaluate("A1")     alert(trigVariable)     alert(firstCellInSheet) }  /* 此示例将工作表 Sheet1 上 A1 单元格的字体设置为加粗。 */ function test() {     Application.Worksheets.Item("Sheet1").Activate()     let boldCell = "A1"     Application.Evaluate(boldCell).Font.Bold = true } ` |

#### **Application.Goto**

选定任意工作簿中的任意区域，并且如果该工作簿未处于活动状态，就激活该工作簿。

**语法**

**express.Goto(Reference, Scroll)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Reference* | 可选          | **Variant**  | 目标。可以是 Range 对象、包含 R1C1-样式批注的单元格引用的字符串。如果省略该参数，目标将为最近一次用 Goto 方法选定的区域。 |
| *Scroll*    | 可选          | **Boolean**  | 如果为 True，则滚动窗口直至区域的左上角出现在窗口的左上角中。如果为 False，则不滚动窗口。默认值为 False。 |

**说明**

该方法与 Select 方法的区别：

1. 如果指定的区域不在位于最前面屏幕的工作表中，ET 将在选定该区域之前切换至该工作表。（如果对不在屏幕的最前面的工作表中的区域使用 Select 方法，则选定该区域时并不激活该工作表）
2. 该方法具有让用户滚动目标窗口的 Scroll 参数
3. 当使用 Goto 方法时，前一次选定区域（Goto 方法运行前）被增加到以前选定区域的数组中（有关详细信息，请参阅 PreviousSelections 属性）。可以使用该功能快速跳过选定区域，选定区域最多为四个
4. Select 方法具有 Replace 参数，而 Goto 方法没有该参数

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例选定工作表 Sheet1 中的单元格 A154，并滚动工作表以显示该区域。 */ function test() {     let rg = Application.ActiveWorkbook.Worksheets.Item("Sheet1").Range("A154");     Application.Goto(rg, true) }` |

#### **Application.InputBox**

显示一个接收用户输入的对话框。返回此对话框中输入的信息。

**语法**

**express.InputBox(Prompt, Title, Default, Left, Top, HelpFile, HelpContextID, Type)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Prompt*        | 必选          | **String**   | 要在对话框中显示的消息。可为字符串、数字、日期、或布尔值（在显示之前，ET 自动将其值强制转换为 String）。 |
| *Title*         | 可选          | **Variant**  | 输入框的标题。如果省略该参数，默认标题将为“Input”。          |
| *Default*       | 可选          | **Variant**  | 指定一个初始值，该值在对话框最初显示时出现在文本框中。如果省略该参数，文本框将为空。该值可以是 Range 对象。 |
| *Left*          | 可选          | **Variant**  | 指定对话框相对于屏幕左上角的 X 坐标（以磅为单位）。          |
| *Top*           | 可选          | **Variant**  | 指定对话框相对于屏幕左上角的 Y 坐标（以磅为单位）。          |
| *HelpFile*      | 可选          | **Variant**  | 此输入框使用的帮助文件名。如果存在 HelpFile 和 HelpContextID 参数，对话框中将出现一个帮助按钮。 |
| *HelpContextID* | 可选          | **Variant**  | HelpFile 中帮助主题的上下文 ID 号。                          |
| *Type*          | 可选          | **Variant**  | 指定返回的数据类型。如果省略该参数，对话框将返回文本。       |

**说明**

下表列出了可以在 Type 参数中传递的值。可以为下列值之一或其中几个值的和。例如，对于一个可接受文本和数字的输入框，将 *Type* 设置为 1 + 2。

| 值   | 含义                                |
| ---- | ----------------------------------- |
| 0    | 公式                                |
| 1    | 数字                                |
| 2    | 文本（字符串）                      |
| 4    | 逻辑值（**True** 或 **False**）     |
| 8    | 单元格引用，作为一个 **Range** 对象 |
| 16   | 错误值，如 #N/A                     |
| 64   | 数值数组                            |

使用 **InputBox** 可以显示一个简单的对话框，以便可以输入要在宏中使用的信息。此对话框有一个“确定”按钮和一个“取消”按钮。如果选择了“确定”按钮，则 **InputBox** 将返回对话框中输入的值。如果单击“取消”按钮，则 **InputBox** 返回 **False**。

如果 *Type* 为 0，**InputBox** 将以文本格式返回公式。例如，“=2*PI()/360”。如果公式中有引用，将以 A1-样式引用返回（使用 **ConvertFormula**转换引用样式）。

如果 *Type* 为 8，**InputBox** 将返回一个 **Range** 对象。您必须用 **Set** 语句将结果指定给一个 **Range** 对象，如下例所示。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myRange = Application.InputBox("Sample", undefined, undefined, undefined, undefined, undefined, undefined,8) }` |

如果不使用 **Set** 语句，此变量将被设置为这个区域的值，而不是 **Range** 这个对象本身。

如果使用 **InputBox** 方法要求用户输入公式，则必须使用**FormulaLocal**属性来将此公式指定给一个 **Range** 对象。输入的公式使用用户语言。

**InputBox** 方法与 **InputBox** 函数的区别在于：它可以对用户的输入进行选择性验证，也可用于 ET 对象、误差值、和公式的输入。注意，`Application.InputBox` 调用的是 **InputBox** 方法，不带对象识别符的 `InputBox` 调用的是 **InputBox** 函数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例提示用户输入数字。 */ function test() {     let myNum = Application.InputBox("Enter a number")  }  /* 本示例提示用户在 Sheet1 中选取一个单元格。示例使用 Type 参数证实返回值是有效的单元格引用 (一个 Range 对象)。 */ function test() {     Application.Worksheets.Item("Sheet1").Activate()     let myCell = Application.InputBox("Select a cell", undefined, undefined, undefined, undefined, undefined, undefined, 8) }` |

#### **Application.Quit**

退出ET

**语法**

**express.Quit()**

*express*   一个代表 **Application** 对象的变量。

**说明**

使用此方法时，如果未保存的工作簿处于打开状态，则 ET 将显示一个对话框，询问是否要保存所作更改。要防止发生这种情况，请在使用 **Quit** 方法前保存所有工作簿或将**DisplayAlerts**属性设置为 **False**。如果该属性为 **False**，则在 ET 退出时，即使有未保存的工作簿，也不会显示对话框，而且不保存就退出。

如果将工作簿的**Saved**属性设置为 **True**，而没有将工作簿保存到磁盘上，则 ET 在退出时不会提示保存该工作簿。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例保存所有打开的工作簿，然后退出 ET。 */ function test() {     let w = Application.Workbooks     for (let i = 1; i <= w.Count; i++) {         w.Item(i).Save()     }     Application.Quit() }` |

#### **Application.Run**

运行一个宏或者调用一个函数。该方法可用于运行宏编辑器编写的宏，或者运行 DLL 或 XLL 中的函数。

**语法**

**express.Run(Macro, Arg1-Arg30)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *Macro*      | 可选          | **Variant**  | 要运行的宏。它可以是具有宏名称的字符串、表示函数所在位置的 Range 对象，或者是一个已注册的 DLL (XLL) 函数的注册号。如果使用字符串，将在当前工作表的上下文中对该字符串求值。 |
| *Arg1-Arg30* | 可选          | **Variant**  | 应传递给函数的参数。                                         |

**说明**

此方法不可使用命名参数，参数必须通过位置进行传递。

**Run** 方法返回被调用的宏返回的任何值。如果将对象作为参数传递给宏，该对象将转换为相应的值（通过对该对象应用 **Value** 属性）。这意味着不能用 **Run** 方法将对象传递给宏。

**成员属性**

#### **Application.ActiveCell**

返回一个 Range 对象，该对象代表活动窗口（顶部窗口）或指定窗口中的活动单元格。如果窗口中没有显示工作表，此属性无效。只读。

**语法**

**express.ActiveCell**

*express*   一个代表 **Application** 对象的变量。

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
| `/* 此示例在消息框中显示活动单元格的值。由于如果活动表不是工作表则 ActiveCell 属性无效，所以此示例使用 ActiveCell 属性之前先激活 Sheet1。 */ function test() {     Application.Worksheets.Item("Sheet1").Activate()     alert(Application.ActiveCell.Value2) }  /* 此示例更改活动单元格的字体格式设置。 */ function test() {     Application.Worksheets.Item("Sheet1").Activate()     let font = Application.ActiveCell.Font     font.Bold = true     font.Italic = true }` |

#### **Application.ActiveSheet**

返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 **null**。

**语法**

**express.ActiveSheet**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果不指定对象识别符，则此属性返回活动工作簿中的活动工作表。

如果某个工作簿出现在若干个窗口中，那么该工作簿的 **ActiveSheet** 属性在不同窗口中可能不同。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例显示活动工作表的名称。 */ function test() {     alert("The name of the active sheet is " + Application.ActiveSheet.Name) }` |

#### **Application.ActiveWindow**

返回一个 **Window**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Window/Window%20.htm#jsObject_Window)对象，该对象代表活动窗口（顶部窗口）。只读。如果没有打开的窗口，则返回 **null**。

**语法**

**express.ActiveWindow**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示活动窗口的名称（Caption 属性）。 */ function test() {     alert("The name of the active window is " + Application.ActiveWindow.Caption) }` |

#### **Application.ActiveWorkbook**

返回一个 **Workbook**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)对象，该对象代表活动窗口（顶部窗口）中的工作簿。只读。如果没有打开的窗口，或者“信息”窗口或“剪贴板”窗口为活动窗口，则返回 **null**。

**语法**

**express.ActiveWorkbook**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示活动工作簿的名称。 */ function test() {     alert("The name of the active workbook is " + Application.ActiveWorkbook.Name) }` |

#### **Application.Application**

该属性返回一个代表 ET 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息。 */ function test() {     let myObject = Application     alert(myObject.Application.Value) }` |

#### **Application.Build**

返回 ET 内部版本号。**Number** 类型，只读。

**语法**

**express.Build**

*express*   一个代表 **Application** 对象的变量。

**说明**

通常检测[ Version ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#Application.Version)属性较为安全，除非必须要获知内部版本号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例检测 Build 属性。 */ function test() {     if(Application.Build > 2500) {         alert("build-dependent code here")     } }` |

#### **Application.Caption**

返回或设置一个 **String** 值，它代表出现在 ET 主窗口标题栏中显示的名称。

**语法**

**express.Caption**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果未设置该名称，或将其设置为 **Empty**, 则此属性返回“ET”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例将出现在 ET 主窗口标题栏中的名称设置为自定义的名称。 */ function test() {     Application.Caption = "KingSoft" }` |

#### **Application.Cells**

返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表活动工作表中的所有单元格。如果活动文档不是工作表，则此属性无效。

**语法**

**express.Cells**

*express*   一个代表 **Application** 对象的变量。

**说明**

因为 **Item** 属性是 [**Range** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象的默认属性，所以可以在 **Cells** 关键字后面紧接着指定行和列索引。有关详细信息，请参阅 **Item** 属性和本主题的示例。

在不使用对象识别符的情况下，使用此属性将返回一个 [**Range** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表活动工作表中所有的单元格。

#### **Application.Columns**

返回一个 [Range](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表活动工作表中的所有列。如果活动文档不是工作表，则 Columns 属性无效。

**语法**

**express.Columns**

*express*   一个代表 **Application** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 **ActiveSheet.Columns**``。

此属性在应用于一个是多重选定区域的 [**Range** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象时，会只从该区域的第一个子区域中返回列。例如，如果 [**Range** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象有两个子区域 A1:B2 和 C3:D4，那么，`Selection.Columns.Count` 的返回值是 2，而不是 4。若要对一个可能包含多重选定区域的区域使用此属性，请测试 `Areas.Count` 以确定此区域内是否包含多个子区域。如果包含，请对此区域内的每个子区域进行循环。

\````

#### **Application.Dialogs**

返回一个 **Dialogs**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Dialogs/Dialogs%20.htm#jsObject_Dialogs)集合，该集合代表所有内置对话框。只读。

**语法**

**express.Dialogs**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示“文件”菜单的“打开”对话框。 */ function test() {     Application.Dialogs.Item(xlDialogOpen).Show() }` |

#### **Application.Height**

返回一个 **Double** 值，它代表主应用程序窗口的高度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果窗口处于最小化，则此属性为只读，并引用图标的高度。如果窗口处于最大化，则无法设置此属性。使用 **WindowState**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#Application.WindowState)属性可确定窗口的状态。

#### **Application.Hwnd**

返回一个 **Number** 类型的值，该值表示 ET 窗口的最高级别的窗口句柄。只读。

**语法**

**express.Hwnd**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 在本示例中，ET 向用户通知 ET 窗口的最高级别的窗口句柄。 */ function test() {     MsgBox("The top-level window handle is: " + Application.Hwnd) }` |

#### **Application.Left**

返回或设置一个 **Double** 值，它代表从屏幕左边缘到 ET 主窗口左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **Application** 对象的变量。

#### **Application.Name**

返回一个代表对象名称的 **String** 值。

**语法**

**express.Name**

*express*   一个代表 **Application** 对象的变量。

#### **Application.Names**

返回一个 **Names**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Names/Names%20.htm#jsObject_Names)集合，该集合代表活动工作簿中的所有名称。只读 **Names**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Names/Names%20.htm#jsObject_Names)对象。

**语法**

**express.Names**

*express*   一个代表 **Application** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveWorkbook.Names`。

#### **Application.NewWorkbook**

返回一个 **NewFile** 对象。

**语法**

**express.NewWorkbook**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 在本示例中，ET 将 wkbOne 变量设置为 NewFile 对象。 */ function test() {     // Create a reference to an instance of the NewFile object.     let wkbOne = Application.NewWorkbook }` |

#### **Application.Path**

返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。

**语法**

**express.Path**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例显示 ET 的完整路径。 */ function test() {     alert("The path is " + Application.Path) }` |

#### **Application.Rows**

返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表活动工作表中的所有行。如果活动文档不是工作表，则 **Rows** 属性失效。**Range** 对象，只读。

**语法**

**express.Rows**

*express*   一个代表 **Application** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveSheet.Rows`。

该属性在应用于是多个选定区域的 **Range** 对象时，只从该区域中第一个子区域内返回行。例如，如果 **Range** 对象有两个子区域：A1:B2 和 C3:D4，则 `Selection.Rows.Count` 返回 2 而不是 4。若要在一个可能包含多个选定区域的区域中使用此属性，请测试 `Areas.Count` 以确定该区域是否包含多个选择区域。如果是，请对此区域内的每个子区域进行循环，如第 3 个示例所示。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例删除 Sheet1 的第三行。 */ function test() {     Application.Worksheets.Item("Sheet1").Rows.Item(3).Delete() }  /* 此示例检查第一张工作表上当前区域中的行，如果某行的第一个单元格值与前一行的第一个单元格的值相等，则删除此行。 */ function test() {     let last;     let cnt = Application.Selection.Rows.Count;     let delCnt = 0;     for (i = 1; i <= cnt - delCnt; i++) {         let rw = Application.Selection.Rows.Item(i);         cur = rw.Cells.Item(1, 1).Value();         if (cur === last) {             rw.Delete();             i--;             delCnt ++;         }         last = cur;     } }   /* 此示例显示 Sheet1 选定区域中的行数。如果选择了多个子区域，此示例将对每一个子区域进行循环。 */ function test() {     Application.Worksheets.Item("Sheet1").Activate()     let areaCount = Application.Selection.Areas.Count     if (areaCount <= 1) {         alert(`The selection contains${Application.Selection.Rows.Count}rows.`)     } else {         let i = 1         for (item in Application.Selection.Areas) {             alert(`Area${i}of the selection contains${item.Rows.Count}rows.`)             i += 1         }     } }` |

#### **Application.Selection**

为 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象返回在活动窗口中选定的对象。

**语法**

**express.Selection**

*express*   一个代表 **Application** 对象的变量。

**说明**

返回的对象类型取决于当前所选内容（例如，如果选择了单元格，此属性将返回 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象）。如果未选择任何内容，**Selection** 属性将返回 **Nothing**。

在不使用对象识别符的情况下，使用此属性等效于使用 `Application.Selection`。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例清空 Sheet1 的选定对象（假定选定对象为单元格区域）。 */ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.Selection.Clear() }  /* 本示例显示选定对象的 wpsjs 加载项 对象类型。 */ function test() {     Application.Worksheets.Item("Sheet1").Activate()     alert("The selection object type is " + typeof(Application.Selection)) }` |

#### **Application.Top**

返回或设置一个 **Double** 值，它代表从屏幕上边缘到 ET 主窗口上边缘的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果应用程序窗口已最小化，则此属性将控制窗口图标的位置（在屏幕的任何位置）。

#### **Application.Value**

返回一个代表应用程序名称的 **String** 值。

**语法**

**express.Value**

*express*   一个代表 **Application** 对象的变量。

**说明**

此属性始终返回 "Microsoft Excel"。

#### **Application.Version**

返回一个 **String** 值，它代表 ET 版本号。

**语法**

**express.Version**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例显示包含 ET 版本号和操作系统名称的消息框。 */ function test() {     alert(`Welcome to ET version ${Application.Version} running on ${Application.OperatingSystem}!`) }` |

#### **Application.Width**

返回或设置一 **个 Double** 值，它代表应用程序窗口的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果窗口处于最小化状态，则 **Width** 为只读，并返回窗口图标的宽度。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*  此示例将活动窗口的尺寸扩大为可用的最大尺寸（假定该窗口尚未最大化）。 */ function test() {     let window = Application.ActiveWindow     window.WindowState = Application.Enum.xlNarrow     window.Top = 1     window.Left = 1     window.Height = Application.UsableHeight     window.Width = Application.UsableWidth }` |

#### **Application.WindowState**

返回或设置窗口的状态。[XlWindowState ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlWindowState%20%E6%9E%9A%E4%B8%BE.html)类型，可读写。

**语法**

**express.WindowState**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将 ET 应用程序窗口最大化。 */ function test() {     Application.WindowState = Application.Enum.xlMaximized }  /*  此示例将活动窗口的尺寸扩大为可用的最大尺寸（假定该窗口尚未最大化）。 */ function test() {     let window = Application.ActiveWindow     window.WindowState = Application.Enum.xlNarrow     window.Top = 1     window.Left = 1     window.Height = Application.UsableHeight     window.Width = Application.UsableWidth }` |

#### **Application.Windows**

返回一个 **Windows** 集合，它代表所有工作簿中的所有窗口。**Windows**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Windows/Windows%20.htm#jsObject_Windows)对象，只读。

**语法**

**express.Windows**

*express*   一个代表 **Application** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 `Application.Windows`。

此属性返回的集合中既包括可见窗口，也包括隐藏窗口。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例关闭 ET 中第一个打开或隐藏的窗口。 */ function test() {     Application.Windows.Item(1).Close() }  /* 此示例将活动工作簿的窗口一命名为“Consolidated Balance Sheet”。此名称将被用作对 Windows 集合的索引。 */ function test() {     Application.ActiveWorkbook.Windows.Item(1).Caption = "Consolidated Balance Sheet"     Application.ActiveWorkbook.Windows.Item("Consolidated Balance Sheet").ActiveSheet.Calculate() }` |

#### **Application.Workbooks**

返回一个 **Workbooks**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbooks/Workbooks%20.htm#jsObject_Workbooks)集合，该集合表示所有打开的工作簿。只读。

**语法**

**express.Workbooks**

*express*   一个代表 **Application** 对象的变量。

**说明**

在不使用对象识别符的情况下，使用该属性相当于使用。

| 示例代码复制            |
| ----------------------- |
| `Application.Workbooks` |

由 **Workbooks** 属性返回的集合并不包含打开的加载宏（一种特殊的隐藏工作簿）。但如果知道文件名，则可返回单个打开的加载宏。例如，

| 示例代码复制                              |
| ----------------------------------------- |
| `Application.Workbooks.Item("Oscar.xla")` |

将把名为“Oscar.xla”的打开的加载宏作为 

Workbook

 对象返回。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例激活 Book1.xls 工作簿。 */ function test() {     Application.Workbooks.Item("Book1").Activate() }  /* 本示例打开 Large.xlsx 工作簿。 */ function test() {     Application.Workbooks.Open("C:/Users/wps/Desktop/large.xlsx") }  /* 本示例关闭除正在运行本示例的工作簿以外的所有其他工作簿，并保存其更改。 */ function test() {     const Count = Application.Workbooks.Count     for(let i = 1; i <= Count; i++){         if(Application.Workbooks.Item(i).Name != Application.ThisWorkbook.Name){             Application.Workbooks.Item(i).Close(true)         }     }   }` |

#### **Application.WorksheetFunction**

返回**WorksheetFunction**对象。只读。

**语法**

**express.WorksheetFunction**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例对单元格区域 A1:A10 应用 Min 工作表函数，并显示结果。 */ function test() {     let myRange = Application.Worksheets.Item("Sheet1").Range("A1:C10")     let answer = Application.WorksheetFunction.Min(myRange)     alert(answer) }` |

#### **Application.Worksheets**

对于 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，返回一个 [Sheets ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)集合，它代表活动工作簿中的所有工作表。对于 **Workbook**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)对象，返回一个 [Sheets ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)集合，它代表指定工作簿中的所有工作表。[Sheets](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)对象，只读。

**语法**

**express.Worksheets**

*express*   一个代表 **Application** 对象的变量。

**说明**

在不使用对象识别符的情况下，使用此属性将返回活动工作簿中所有的工作表。

此属性不返回宏表；使用 **Excel4MacroSheets** 属性或 **Excel4IntlMacroSheets** 属性可返回这些表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例显示活动工作簿中 sheet1 上单元格 A1 中的值。。 */ function test() {     alert(Application.Worksheets.Item("Sheet1").Range('A1').Value2) }  /* 此示例显示活动工作簿中每个工作表的名称。 */ function test() {     for (let i = 1; i <= Application.Worksheets.Count; i++) {         alert(Application.Worksheets.Item(i).Name)     } }  /* 此示例向活动工作簿添加新工作表，并设置该工作表的名称。 */ function test() {     let newSheet = Application.Worksheets.Add()     newSheet.Name = "current Budget" }` |

适用环境：web

适用平台：windows/linux