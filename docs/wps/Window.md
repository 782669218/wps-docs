#### **Window**



代表一个窗口。许多文档特征（如滚动条和标尺）实际上是窗口的属性。

**说明**

**Window** 对象是 **Windows** 集合的一个成员。**Application** 对象的 **Windows** 集合包含应用程序中的所有窗口，而 **Document** 对象的 **Windows** 集合仅包含显示指定文档的窗口。

使用 **Windows**(*Index*) 可返回一个 **Window** 对象，其中 *Index* 为窗口名称或索引号。以下示例将 Document1 窗口最大化。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Windows.Item("Document1").WindowState = wdWindowStateMaximize` |

索引号是

“窗口”

菜单中窗口名称左侧的数字。以下示例显示 

Windows

 集合中第一个窗口的标题。

| 示例代码复制                     |
| -------------------------------- |
| `alert(Windows.Item(1).Caption)` |

使用 

Add

 方法或 

NewWindow

 方法可向 

Windows

 集合添加新窗口。下列每个语句都为活动窗口中的文档创建一个新窗口。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.ActiveWindow.NewWindow() Application.Windows.Add()` |

对同一文档打开多个窗口时，窗口标题中将出现一个冒号 (:) 和一个数字。

在将视图切换到打印预览时，会创建一个新窗口。关闭打印预览时，该窗口从 **Windows** 集合中删除。

**方法**

|                                                              | 名称                       | 说明                                                         |
| ------------------------------------------------------------ | -------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**               | 激活指定窗口。                                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Close**                  | 关闭指定的窗口。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetPoint**               |                                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LargeScroll**            | 将窗口或窗格滚动指定的屏数。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NewWindow**              | 打开一个新窗口，该窗口与指定窗口显示同一个文档。返回一个 **Window** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PageScroll**             | 使指定的窗格或窗口逐页滚动。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintOut**               | 打印指定窗口中显示的全部或部分文档。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RangeFromPoint**         | 返回 **Range** 或 **Shape** 对象，该对象位于由屏幕位置坐标对指定的位置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScrollIntoView**         | 滚动文档窗口，以便在文档窗口显示指定的区域或图形。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetFocus**               | 在指定的文档窗口为电子邮件设置焦点。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SmallScroll**            | 将窗口或窗格滚动指定的行数。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ToggleRibbon**           | 显示或隐藏功能区。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ToggleShowAllReviewers** | 显示或隐藏包含备注和修订的文档中的所有备注。                 |

**属性**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Active**                     | 如果指定窗口处于活动状态，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActivePane**                 | 返回 **Pane** 对象，该对象代表指定窗口的活动窗格。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                | 返回一个代表 WPS Office WPS 应用程序的 **Application** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**                    | 返回或设置显示在文档或应用程序窗口标题栏中的窗口题注文字。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                    | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayHorizontalScrollBar** | 如果在指定的窗口中显示水平滚动条，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayLeftScrollBar**       | 如果垂直滚动条显示在文档窗口左侧，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayRightRuler**          | 如果垂直标尺显示在页面视图中文档窗口右侧，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayRulers**              | 如果在指定的窗口或窗格中显示标尺，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayScreenTips**          | 如果批注、脚注、尾注和超链接以提示形式显示，则该属性值为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayVerticalRuler**       | 如果在指定的窗口或窗格中显示垂直标尺，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Document**                   | 返回一个 **Document** 对象，该对象与指定窗格、窗口或选定内容相关。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DocumentMap**                | 如果显示文档结构图，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnvelopeVisible**            | 如果电子邮件的页眉在文档窗口是可见的，则该属性值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**                     | 返回或设置窗口的高度。Long 类型，可读写。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalPercentScrolled**  | 返回或设置水平滚动条位置（以文档宽度的百分比表示）。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IMEMode**                    | 返回或设置日文输入法编辑器 (IME) 的默认启动模式。**WdIMEMode** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**                      | 返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                       | 返回或设置一个 **Long** 类型的值，该值代表指定窗口的水平位置，以磅为单位。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Next**                       | 返回打开的文档窗口集合中的下一个文档窗口。只读。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Panes**                      | 返回一个 **Panes** 集合，该集合代表指定窗口中所有的窗格。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                     | 返回一个 **Object** 类型值，该值代表指定 **Window** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Previous**                   | 返回打开的文档窗口集合中的上一个文档窗口。只读。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Selection**                  | 返回一个 **Selection** 对象，该对象代表一个选定范围或插入点。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowSourceDocuments**        | 返回或设置一个 **WdShowSourceDocuments** 常量，该常量代表 WPS 在比较和合并过程之后显示源文档的方式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Split**                      | 如果将窗口拆分为多个窗格，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SplitVertical**              | 返回或设置指定窗口的垂直拆分百分比。**Long** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StyleAreaWidth**             | 该属性返回或设置样式区的宽度，以磅为单位。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Thumbnails**                 | 设置或返回一个 **Boolean** 类型的值，代表是否沿 WPS 文档窗口的左侧显示文档页面的缩略图。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                        | 返回或设置指定文档窗口的垂直位置（以磅为单位）。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                       | 返回窗口的类型。只读 **WdWindowType** 类型。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UsableHeight**               | 返回指定的文档窗口中活动工作区的高度（以磅为单位）。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UsableWidth**                | 返回指定的文档窗口中活动工作区的宽度（以磅为单位）。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalPercentScrolled**    | 返回或设置垂直滚动条的位置（以文档长度的百分比表示）。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **View**                       | 返回一个 **View** 对象，该对象代表指定窗口或窗格的视图。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**                    | 如果指定对象可见，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**                      | 返回或设置指定文档窗口的宽度（以磅为单位）。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WindowNumber**               | 该属性返回在指定窗口中显示的文档窗口序号。例如，如果窗口标题为“Sales.doc:2”，则该属性返回数值 2。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WindowState**                | 返回或设置指定文档窗口或任务窗口的状态。**WdWindowState** 类型，可读写。 |

**成员方法**

#### **Window.Activate**

激活指定窗口。

**语法**

**express.Activate()**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例激活 Windows 集合中的下一个窗口。*/ function NextWindow() {     //Two or more documents must be open for this statement to execute.     ActiveDocument.ActiveWindow.Next.Activate() }` |

#### **Window.Close**

关闭指定的窗口。

**语法**

**express.Close(SaveChanges, RouteDocument)**

*express*   一个代表 **Window** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SaveChanges*   | 可选          | **Variant**  | 指定保存文档的操作。可以是下列 WdSaveOptions 常量之一：wdDoNotSaveChanges、wdPromptToSaveChanges 或 wdSaveChanges。 |
| *RouteDocument* | 可选          | **Variant**  | 如果该参数值为 True，则将文档传送给下一个收件人。如果文档没有附加的传送名单，则忽略此参数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭活动文档的活动窗口并保存文档。*/ ActiveDocument.ActiveWindow.Close(wdSaveChanges)` |

#### **Window.GetPoint**

**语法**

**express.GetPoint(ScreenPixelsLeft, ScreenPixelsTop, ScreenPixelsWidth, ScreenPixelsHeight, obj)**

*express*   一个代表 **Window** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**                                  |
| -------------------- | ------------- | ------------ | ----------------------------------------- |
| *ScreenPixelsLeft*   | 必选          | **Long**     | 需要 WPS 为对象的左边界值返回的变量名。   |
| *ScreenPixelsTop*    | 必选          | **Long**     | 需要 WPS 为对象的顶部边界值返回的变量名。 |
| *ScreenPixelsWidth*  | 必选          | **Long**     | 需要 WPS 为对象的宽度返回的变量名。       |
| *ScreenPixelsHeight* | 必选          | **Long**     | 需要 WPS 为对象的高度值返回的变量名。     |
| *obj*                | 必选          | **Object**   | Range 或 Shape 对象。                     |

**说明**

如果在屏幕上看不到整个区域或图形，则导致出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查当前选定对象并返回其屏幕坐标。*/ function test() { let pLeft=0.1 let pTop=0.1 let pWidth=0.1 let pHeight=0.1 ActiveWindow.GetPoint(pLeft, pTop, pWidth, pHeight, Selection.Range) MsgBox("Left = " + pLeft + "\n" +         "Top = " + pTop + "\n" +         "Width = " + pWidth + "\n" +         "Height = " + pHeight) }` |

#### **Window.LargeScroll**

将窗口或窗格滚动指定的屏数。

**语法**

**express.LargeScroll(Down, Up, ToRight, ToLeft)**

*express*   一个代表 **Window** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**             |
| --------- | ------------- | ------------ | -------------------- |
| *Down*    | 可选          | **Variant**  | 向下滚动窗口的屏数。 |
| *Up*      | 可选          | **Variant**  | 向上滚动窗口的屏数。 |
| *ToRight* | 可选          | **Variant**  | 向右滚动窗口的屏数。 |
| *ToLeft*  | 可选          | **Variant**  | 向左滚动窗口的屏数。 |

**说明**

此方法等效于在水平和垂直滚动条上单击滚动块以外的位置。

如果指定了 *Down* 和 *Up* 两个参数，则窗口滚动屏数为这两个参数的差。例如，如果 *Down* 为 2，*Up* 为 4，则窗口向上滚动两屏。同样，如果指定了 *ToLeft* 和 *ToRight* 两个参数，则窗口滚动屏数也为这两个参数的差。

这四个参数都可以取负值。如果没有指定参数，窗口将向下滚动一屏。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动窗口下滚一屏。*/ ActiveDocument.ActiveWindow.LargeScroll(1)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例拆分活动窗口，并且上滚两屏，右滚一屏。*/ function test() { ActiveDocument.ActiveWindow.Split = true ActiveDocument.ActiveWindow.LargeScroll(null, 2, 1) }` |

#### **Window.NewWindow**

打开一个新窗口，该窗口与指定窗口显示同一个文档。返回一个 **Window** 对象。

**语法**

**express.NewWindow()**

*express*   一个代表 **Window** 对象的变量。

**返回值**

Window

**说明**

当多个窗口打开同一个文档时，将在窗口标题中显示一个冒号 (:) 和一个数字。下列两条指令在功能上是等效的。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let myWindow = ActiveDocument.ActiveWindow.NewWindow() let myWindow = NewWindow() }` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在为 Document1 打开一个新窗口之前和之后各发布一条消息，以指示现存窗口的数目。*/ function test() { MsgBox(Windows.Count + " windows open") Windows.Item("Document1").NewWindow() MsgBox(Windows.Count + " windows open") }` |

#### **Window.PageScroll**

使指定的窗格或窗口逐页滚动。

**语法**

**express.PageScroll(Down, Up)**

*express*   一个代表 **Window** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                           |
| -------- | ------------- | ------------ | -------------------------------------------------- |
| *Down*   | 可选          | **Variant**  | 向下滚动的页数。如果省略此参数，则参数值设置为 1。 |
| *Up*     | 可选          | **Variant**  | 向上滚动的页数。                                   |

**说明**

**PageScroll** 方法仅用于页面视图或 Web 版式视图。此方法不影响插入点的位置。

如果指定了 *Down* 和 *Up* 两个参数，则窗口滚动页数为这两个参数的差。例如，如果 *Down* 为 2，*Up* 为 4，则窗口向上滚动两页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例使活动窗口向下滚动三页。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdPrintView ActiveDocument.ActiveWindow.PageScroll(3) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例使活动窗口向下滚动一页。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdPrintView ActiveDocument.ActiveWindow.PageScroll() }` |

#### **Window.PrintOut**

打印指定窗口中显示的全部或部分文档。

**语法**

**express.PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)**

*express*   一个代表 **Window** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Background*           | 可选          | **Variant**  | 如果将该属性设置为 True，则 WPS 在打印文档时继续运行宏。     |
| *Append*               | 可选          | **Variant**  | 如果将该属性设置为 True，则将指定文档附加到 OutputFileName 参数所指定的文件名。如果将该属性设置为 False，则覆盖 OutputFileName 参数所指定文件的内容。 |
| *Range*                | 可选          | **Variant**  | 页码范围。可以是任意 WdPrintOutRange 常量。                  |
| *OutputFileName*       | 可选          | **Variant**  | 如果 PrintToFile 为 True，则该参数指定输出文件的路径和文件名。 |
| *From*                 | 可选          | **Variant**  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定起始页码。    |
| *To*                   | 可选          | **Variant**  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定结束页码。    |
| *Item*                 | 可选          | **Variant**  | 要打印的项目。可以是任意 WdPrintOutItem 常量。               |
| *Copies*               | 可选          | **Variant**  | 要打印的份数。                                               |
| *Pages*                | 可选          | **Variant**  | 要打印的页码和页码范围，中间用逗号分开。例如，“2, 6-10”表示打印第 2 页以及第 6 至第 10 页。 |
| *PageType*             | 可选          | **Variant**  | 要打印的页面类型。可以是任意 WdPrintOutPages 常量。          |
| *PrintToFile*          | 可选          | **Variant**  | 如果该属性值为 True，则将打印指令发送到文件。请确保使用 OutputFileName 指定文件名。 |
| *Collate*              | 可选          | **Variant**  | 在打印文档的多份副本时，如果该属性值为 True，则完成打印所有页面后再打印下一份副本。 |
| *FileName*             | 可选          | **Variant**  | 要打印的文档的路径和文件名。如果省略该参数， WPS 将打印活动文档（仅应用于 Application 对象）。 |
| *ActivePrinterMacGX*   | 可选          | **Variant**  | 该参数仅适用于 WPS OfficeMacintosh Edition。有关该参数的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。 |
| *ManualDuplexPrint*    | 可选          | **Variant**  | 如果该属性值为 True，则在无双面打印组件的打印机上打印双面文档。如果该属性值为 True，则忽略 PrintBackground 和 PrintReverse 属性。使用 PrintOddPagesInAscendingOrder 和 PrintEvenPagesInAscendingOrder 属性可在手动双面打印时控制输出。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *PrintZoomColumn*      | 可选          | **Variant**  | 表示 WPS 在一页纸上水平放置的页数。可以是 1、2、3 或 4 页。与 PrintZoomRow 参数一同使用可在单张纸上打印多页文档。 |
| *PrintZoomRow*         | 可选          | **Variant**  | 表示 WPS 在一页纸上垂直放置的页数。可以是 1、2 或 4 页。与 PrintZoomColumn 参数一同使用可在单张纸上打印多页文档。 |
| *PrintZoomPaperWidth*  | 可选          | **Variant**  | WPS 将打印页面缩放到的宽度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。 |
| *PrintZoomPaperHeight* | 可选          | **Variant**  | WPS 将打印页面缩放到的高度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印活动文档的当前页面。*/ ActiveDocument.PrintOut(null,null,wdPrintCurrentPage)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印当前文件夹中的所有文档。Dir 函数用于返回所有扩展名为“.doc”的文件名。*/ function test() { adoc = Dir("*.DOC") Do While adoc <> ""     Application.PrintOut FileName:=adoc     adoc = Dir() Loop }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印活动窗口中文档的前三页。*/ ActiveDocument.ActiveWindow.PrintOut(null,null,wdPrintFromTo,null,"1","3")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印活动文档中的备注。*/ function test() { if(ActiveDocument.Comments.Count >= 1){     ActiveDocument.PrintOut(null,null,null,null,null,null,wdPrintComments) } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将打印活动文档，每张纸上打印六页文档。*/ ActiveDocument.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,3,2)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例按实际尺寸的 75% 打印活动文档。*/ ActiveDocument.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,0.75 * (8.5 * 1440),0.75 * (11 * 1440))` |

#### **Window.RangeFromPoint**

返回 **Range** 或 **Shape** 对象，该对象位于由屏幕位置坐标对指定的位置。

**语法**

**express.RangeFromPoint(x, y)**

*express*   一个代表 **Window** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                         |
| -------- | ------------- | ------------ | ------------------------------------------------ |
| *x*      | 必选          | **Long**     | 从屏幕左边缘到该点的水平距离（以像素为单位）。   |
| *y*      | 必选          | **Long**     | 从屏幕顶部边缘到该点的垂直距离（以像素为单位）。 |

**返回值**

Object

**说明**

如果在坐标对指定的位置中没有区域或图形，则该方法返回 **Nothing**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例可实现的功能是：新建一个文档，并在其中添加一个五角星。然后便获取该图形的屏幕位置并计算其中心位置。用这些坐标（本示例使用 RangeFromPoint 方法）来返回一个到该图形的引用并且改变其填充颜色。*/ function test() { let pLeft = 0.1 let pTop = 0.1 let pWidth = 0.1 let pHeight = 0.1 let newShape let newDoc = Documents.Add()  newDoc.Shapes.AddShape(msoShape5pointStar, 288, 100, 100, 72) newDoc.ActiveWindow.GetPoint(pLeft, pTop, pWidth, pHeight, newDoc.Shapes.Item(1)) newShape = newDoc.ActiveWindow.RangeFromPoint(pLeft + pWidth * 0.5, pTop + pHeight * 0.5) newShape.Fill.ForeColor.RGB = (80, 160, 130) }` |

#### **Window.ScrollIntoView**

滚动文档窗口，以便在文档窗口显示指定的区域或图形。

**语法**

**express.ScrollIntoView(Obj, Start)**

*express*   一个代表 **Window** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Obj*    | 必选          | **Object**   | **Range** 或 **Shape** 对象。                                |
| *Start*  | 可选          | **Boolean**  | 如果该区域或形状的左上角出现在文档窗口的左上角，则为 **True**。如果该区域或形状的右下角出现在文档窗口的右下角，则为 **False**。默认值为 **True**。 |

**说明**

如果该区域或图形大于文档窗口，*Start* 参数可以指定该区域或图形的显示部分或获得初始焦点的部分。此方法无法在大纲视图中使用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例滚动活动文档，以便可以在文档窗口看到选定内容*/ Application.ActiveWindow.ScrollIntoView(Selection.Range, true)` |

#### **Window.SetFocus**

在指定的文档窗口为电子邮件设置焦点。

**语法**

**express.SetFocus()**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果该文档不是电子邮件，则此方法将不起作用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：显示电子邮件的页眉，并给邮件正文设置焦点。*/ function test() { ActiveWindow.EnvelopeVisible = true ActiveWindow.SetFocus() }` |

#### **Window.SmallScroll**

将窗口或窗格滚动指定的行数。

**语法**

**express.SmallScroll(Down, Up, ToRight, ToLeft)**

*express*   一个代表 **Window** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Down*    | 可选          | **Variant**  | 向下滚动窗口的行数。一“行”对应于单击一次垂直滚动条上的向下滚动箭头所滚动的距离。 |
| *Up*      | 可选          | **Variant**  | 向上滚动窗口的行数。一“行”对应于单击一次垂直滚动条上的向上滚动箭头所滚动的距离。 |
| *ToRight* | 可选          | **Variant**  | 向右滚动窗口的行数。一“行”对应于单击一次水平滚动条上的向右滚动箭头所滚动的距离。 |
| *ToLeft*  | 可选          | **Variant**  | 向左滚动窗口的行数。一“行”对应于单击一次水平滚动条上的向左滚动箭头所滚动的距离。 |

**说明**

此方法等效于单击水平和垂直滚动条上的滚动箭头。

如果指定了 *Down* 和 *Up* 两个参数，则窗口滚动行数为这两个参数的差。例如，如果 *Down* 为 3，*Up* 为 6，则窗口将向上滚动三行。同样，如果指定了 *ToLeft* 和 *ToRight* 两个参数，则窗口滚动行数也为这两个参数的差。

这四个参数都可以取负值。如果没有指定参数，窗口将下滚一行。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动窗口下滚一行。*/ ActiveDocument.ActiveWindow.SmallScroll(1)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例拆分活动窗口，并上滚及左滚窗口。*/ function test() { ActiveDocument.ActiveWindow.Split = true ActiveDocument.ActiveWindow.SmallScroll(null, 5, null, 5) }` |

#### **Window.ToggleRibbon**

显示或隐藏功能区。

**语法**

**express.ToggleRibbon()**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果功能区可见，则 **TobbleRibbon** 方法隐藏功能区。如果功能区隐藏，则 **ToggleRibbon** 方法显示功能区。

#### **Window.ToggleShowAllReviewers**

显示或隐藏包含备注和修订的文档中的所有备注。

**语法**

**express.ToggleShowAllReviewers()**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档中的所有备注。本示例假设文档中存在备注。*/ function test() { Application.ActiveWindow.ToggleShowAllReviewers() }` |

**成员属性**

#### **Window.Active**

如果指定窗口处于活动状态，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.Active**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果 Windows 集合中的第一个窗口当前不处于活动状态，则本示例将激活该窗口。*/ function ActiveWin() {     if(Windows.Item(1).Active == fasle){         Windows.Item(1).Activate()     } }` |

#### **Window.ActivePane**

返回 **Pane** 对象，该对象代表指定窗口的活动窗格。只读。

**语法**

**express.ActivePane**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例拆分活动窗口，然后激活活动窗格的下一个窗格。*/ function test() { let win = ActiveDocument.ActiveWindow win.Split = true win.ActivePane.Next.Activate() MsgBox("Pane " + win.ActivePane.Index + " is active") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例激活第一个窗口，并显示活动窗格中的制表符。*/ function test() { Application.Windows.Item(1).Activate() Application.Windows.Item(1).ActivePane.View.ShowTabs = true }  ` |

#### **Window.Application**

返回一个代表 WPS Office WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Window** 对象的变量。

**说明**

可以通过 Visual Basic 的 **CreateObject** 和 **GetObject** 函数从 示例代码 (VBA) 项目访问 OLE 自动化对象。

#### **Window.Caption**

返回或设置显示在文档或应用程序窗口标题栏中的窗口题注文字。**String** 类型，可读写。

**语法**

**express.Caption**

*express*   一个代表 **Window** 对象的变量。

**说明**

要将应用程序窗口的题注改为默认文本，请将该属性设置为空字符串（""）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Windows 集合中每个窗口的题注*/ for(let i=1;i<=Windows.Count;i++){     MsgBox(Windows.Item(i).Caption, "Window" + i + " Caption") }  /*本示例重新设置应用程序窗口的题注。*/ Application.Caption = ""  /*本示例将活动窗口的题注设置为与活动文档的名称相同。*/ Application.ActiveDocument.ActiveWindow.Caption = ActiveDocument.FullName  /*本示例改变 WPS 应用程序窗口的题注，使其包括用户名。*/ Application.Caption = UserName + "'s copy of Word"  /*本示例插入一个“表格”题注，然后将第一个图表目录的题注改为“表格”。*/ Application.Selection.Collapse(wdCollapseStart) Application.Selection.Range.InsertCaption("Table") if(Application.ActiveDocument.TablesOfFigures.Count >= 1){     Application.ActiveDocument.TablesOfFigures.Item(1).Caption = "Table" } ` |

#### **Window.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Window.DisplayHorizontalScrollBar**

如果在指定的窗口中显示水平滚动条，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayHorizontalScrollBar**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动窗口中显示水平和垂直滚动条。*/ function test() { ActiveDocument.ActiveWindow.DisplayHorizontalScrollBar = true ActiveDocument.ActiveWindow.DisplayVerticalScrollBar = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例切换 Document1 窗口中的水平滚动条。*/ function test() { let winTemp = Windows.Item("Document1") winTemp.DisplayHorizontalScrollBar = !winTemp.DisplayHorizontalScrollBar }` |

#### **Window.DisplayLeftScrollBar**

如果垂直滚动条显示在文档窗口左侧，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayLeftScrollBar**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动窗口左侧显示垂直滚动条。*/ ActiveWindow.DisplayLeftScrollBar = true` |

#### **Window.DisplayRightRuler**

如果垂直标尺显示在页面视图中文档窗口右侧，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayRightRuler**

*express*   一个代表 **Window** 对象的变量。

**说明**

有关在 WPS 中使用从右向左语言的详细信息，请参阅 WPS 从右向左的语言功能。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口设为页面视图，并在窗口右侧显示垂直标尺。*/ function test() { ActiveWindow.View = wdPrintView ActiveWindow.DisplayRightRuler = true }` |

#### **Window.DisplayRulers**

如果在指定的窗口或窗格中显示标尺，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayRulers**

*express*   一个代表 **Window** 对象的变量。

**说明**

该属性与**“视图”**菜单中的“标尺”命令等效。如果 **DisplayRulers** 为 **False**，则无论 **DisplayVerticalRuler** 属性的状态如何，都不显示水平和垂直标尺。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例切换活动窗口中的标尺显示。*/ ActiveDocument.ActiveWindow.DisplayRulers = !ActiveDocument.ActiveWindow.DisplayRulers` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将窗口切换到页面视图并显示水平和垂直标尺。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdPrintView ActiveDocument.ActiveWindow.DisplayVerticalRuler = true ActiveDocument.ActiveWindow.DisplayRulers = true }` |

#### **Window.DisplayScreenTips**

如果批注、脚注、尾注和超链接以提示形式显示，则该属性值为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.DisplayScreenTips**

*express*   一个代表 **Window** 对象的变量。

**说明**

突出显示那些标记有批注的文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使 WPS 将批注、脚注和尾注显示为提示。而且突出显示那些标记有批注的文本。*/ Application.DisplayScreenTips = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“ WPS 选项”对话框中“显示”选项卡上“页面显示选项”部分中“悬停时显示文档工具提示”复选框的当前状态。*/ let temp = Application.DisplayScreenTips` |

#### **Window.DisplayVerticalRuler**

如果在指定的窗口或窗格中显示垂直标尺，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayVerticalRuler**

*express*   一个代表 **Window** 对象的变量。

**说明**

垂直标尺只出现在页面视图中，且 **DisplayRulers** 属性必须设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Windows 集合中的各个窗口切换到页面视图并显示水平和垂直标尺。*/ function test() { let winLoop  for(let i=1;i<=Windows.Count;i++){     winLoop = Windows.Item(i)     winLoop.View.Type = wdPrintView     winLoop.DisplayRulers = true     winLoop.DisplayVerticalRuler = true } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏活动窗口的水平和垂直标尺。*/ function test() { ActiveDocument.ActiveWindow.DisplayVerticalRuler = false ActiveDocument.ActiveWindow.DisplayRulers = false }` |

#### **Window.Document**

返回一个 **Document** 对象，该对象与指定窗格、窗口或选定内容相关。只读。

**语法**

**express.Document**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*该示例将 myDoc 设置为与活动窗口相关的文档。然后焦点切换至下一个窗口，并加以拆分。再使用 Activate 方法切换回原文档。*/ function test() { let myDoc = Application.ActiveWindow.Document if(Windows.Count >= 2) {      Application.ActiveWindow.Next.Activate()     Application.ActiveWindow.Split = true     myDoc.Activate() } }` |

#### **Window.DocumentMap**

如果显示文档结构图，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DocumentMap**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例切换活动窗口的文档结构图的显示。*/ ActiveDocument.ActiveWindow.DocumentMap = !ActiveDocument.ActiveWindow.DocumentMap` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Sales.doc 窗口的文档结构图。*/ function test() { let docSales = Documents.Open("C:\\Documents\\Sales.doc") docSales.ActiveWindow.DocumentMap = true }` |

#### **Window.EnvelopeVisible**

如果电子邮件的页眉在文档窗口是可见的，则该属性值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.EnvelopeVisible**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果文档不是电子邮件，则该属性将不起作用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将显示电子邮件的页眉。*/ ActiveWindow.EnvelopeVisible = true` |

#### **Window.Height**

返回或设置窗口的高度。Long 类型，可读写。

**语法**

**express.Height**

*express*   一个代表 **Window** 对象的变量。

**说明**

当窗口处于最小化或最大化时，不能设置该属性。使用 **Application** 对象的 **UsableHeight** 属性可确定窗口的最大尺寸。使用 **WindowState** 属性可确定窗口状态。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对活动窗口的高度加以调整，以铺满应用程序的窗口区域。*/ function test() { ActiveDocument.ActiveWindow.WindowState = wdWindowStateNormal ActiveDocument.ActiveWindow.Height = Application.UsableHeight }` |

#### **Window.HorizontalPercentScrolled**

返回或设置水平滚动条位置（以文档宽度的百分比表示）。可读写 **Long** 类型。

**语法**

**express.HorizontalPercentScrolled**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动窗口水平滚动的百分比。*/ MsgBox(ActiveDocument.ActiveWindow.HorizontalPercentScrolled + "%")` |

#### **Window.IMEMode**

返回或设置日文输入法编辑器 (IME) 的默认启动模式。**WdIMEMode** 类型，可读写。

**语法**

**express.IMEMode**

*express*   一个代表 **Window** 对象的变量。

#### **Window.Index**

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

*express*   一个代表 **Window** 对象的变量。

#### **Window.Left**

返回或设置一个 **Long** 类型的值，该值代表指定窗口的水平位置，以磅为单位。可读写。

**语法**

**express.Left**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动窗口的水平位置设置为 100 磅。*/ function test() { let win = ActiveDocument.ActiveWindow win.WindowState = wdWindowStateNormal win.Left = 100 win.Top = 0 }` |

#### **Window.Next**

返回打开的文档窗口集合中的下一个文档窗口。只读。

**语法**

**express.Next**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例激活下一个窗口。*/ function test() { if(Windows.Count > 1){     ActiveDocument.ActiveWindow.Next.Activate() } }  ` |

#### **Window.Panes**

返回一个 **Panes** 集合，该集合代表指定窗口中所有的窗格。

**语法**

**express.Panes**

*express*   一个代表 **Window** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口拆分为两个窗格。*/ function test() { if(ActiveDocument.ActiveWindow.Panes.Count == 1) {     ActiveDocument.ActiveWindow.Panes.Add() } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例激活 Document2 窗口中的第一个窗格。*/ Windows.Item("Document2").Panes.Item(1).Activate()` |

#### **Window.Parent**

返回一个 **Object** 类型值，该值代表指定 **Window** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Window** 对象的变量。

#### **Window.Previous**

返回打开的文档窗口集合中的上一个文档窗口。只读。

**语法**

**express.Previous**

*express*   一个代表 **Window** 对象的变量。

#### **Window.Selection**

返回一个 **Selection** 对象，该对象代表一个选定范围或插入点。只读。

**语法**

**express.Selection**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将第一个窗口的所选内容复制到下一个窗口。*/ function test() { if(Windows.Count >= 2) {     Windows.Item(1).Selection.Copy()     Windows.Item(1).Next.Activate()     Selection.Paste() } }` |

#### **Window.ShowSourceDocuments**

返回或设置一个 **WdShowSourceDocuments** 常量，该常量代表 WPS 在比较和合并过程之后显示源文档的方式。可读写。

**语法**

**express.ShowSourceDocuments**

*express*   一个代表 **Window** 对象的变量。

#### **Window.Split**

如果将窗口拆分为多个窗格，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Split**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口拆分为大小相等的两个窗格。*/ ActiveDocument.ActiveWindow.Split = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果拆分 Document1 窗口，则本示例关闭活动窗格。*/ function test() { if(Windows.Item("Document1").Split == true) {     Windows.Item("Document1").ActivePane.Close() } }` |

#### **Window.SplitVertical**

返回或设置指定窗口的垂直拆分百分比。**Long** 类型，可读写。

**语法**

**express.SplitVertical**

*express*   一个代表 **Window** 对象的变量。

**说明**

若要取消拆分，可将该属性设置为零 (0) 或将 **Split** 属性设置为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例拆分活动窗口，使顶部窗格占窗口的 70%。*/ ActiveDocument.ActiveWindow.SplitVertical = 70` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例垂直均分 Document1 窗口。*/ Windows.Item("Document1").SplitVertical = 50` |

#### **Window.StyleAreaWidth**

该属性返回或设置样式区的宽度，以磅为单位。**Single** 类型，可读写。

**语法**

**express.StyleAreaWidth**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果 **StyleAreaWidth** 属性值大于 0（零），则样式名显示在文本的左侧。在页面视图或 Web 版式视图中不显示样式区。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口切换至普通视图，并设置样式区宽度为 1 英寸。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdNormalView ActiveDocument.ActiveWindow.StyleAreaWidth = InchesToPoints(1) }  ` |

#### **Window.Thumbnails**

设置或返回一个 **Boolean** 类型的值，代表是否沿 WPS 文档窗口的左侧显示文档页面的缩略图。

**语法**

**express.Thumbnails**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例显示活动文档中页面的缩略图。*/ ActiveDocument.ActiveWindow.Thumbnails = true` |

#### **Window.Top**

返回或设置指定文档窗口的垂直位置（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Top**

*express*   一个代表 **Window** 对象的变量。

#### **Window.Type**

返回窗口的类型。只读 **WdWindowType** 类型。

**语法**

**express.Type**

*express*   一个代表 **Window** 对象的变量。

#### **Window.UsableHeight**

返回指定的文档窗口中活动工作区的高度（以磅为单位）。**Long** 类型，只读。

**语法**

**express.UsableHeight**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果在文档窗口中看不见任何工作区，则 **UsableHeight** 返回 1。实际可用的高度可用 **UsableHeight** 值减去 1 来确定。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档窗口中工作区的大小。*/ function test() { let win = ActiveDocument.ActiveWindow  MsgBox("Working area height = " + win.UsableHeight + "\n" +  	"Working area width = " + win.UsableWidth) }  ` |

#### **Window.UsableWidth**

返回指定的文档窗口中活动工作区的宽度（以磅为单位）。**Long** 类型，只读。

**语法**

**express.UsableWidth**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果在文档窗口中看不见任何工作区，则 **UsableWidth** 返回 1。要确定实际可用的高度，可用 **UsableWidth** 值减去 1。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档窗口中工作区的大小。*/ function test() { let win = ActiveDocument.ActiveWindow  MsgBox("Working area height = " + win.UsableHeight + "\n" +  	"Working area width = " + win.UsableWidth) }` |

#### **Window.VerticalPercentScrolled**

返回或设置垂直滚动条的位置（以文档长度的百分比表示）。可读写 **Long** 类型。

**语法**

**express.VerticalPercentScrolled**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动窗口垂直滚动的百分比。*/ MsgBox(ActiveDocument.ActiveWindow.VerticalPercentScrolled + "%")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动窗口垂直滚动 10%。*/ function test() { let aWindow = ActiveDocument.ActiveWindow aWindow.VerticalPercentScrolled = aWindow.VerticalPercentScrolled + 10 }  ` |

#### **Window.View**

返回一个 **View** 对象，该对象代表指定窗口或窗格的视图。

**语法**

**express.View**

*express*   一个代表 **Window** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动窗口切换为全屏显示。*/ ActiveDocument.ActiveWindow.View.FullScreen = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 Windows 集合中所有窗口的视图选项。*/ function test() { let myWindow  for(let i=1; i<=Windows.Count; i++) {     myWindow = Windows.Item(i)     myWindow.View.ShowTabs = true     myWindow.View.ShowParagraphs = true     myWindow.View.Type = wdNormalView } }  ` |

#### **Window.Visible**

如果指定对象可见，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Visible**

*express*   一个代表 **Window** 对象的变量。

**说明**

如果 **Visible** 属性为 **False**，则某些方法和属性可能无法使用。

#### **Window.Width**

返回或设置指定文档窗口的宽度（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Width**

*express*   一个代表 **Window** 对象的变量。

#### **Window.WindowNumber**

该属性返回在指定窗口中显示的文档窗口序号。例如，如果窗口标题为“Sales.doc:2”，则该属性返回数值 2。**Long** 类型，只读。

**语法**

**express.WindowNumber**

*express*   一个代表 **Window** 对象的变量。

**说明**

使用 属性可返回 **Windows** 集合中指定窗口的序号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检索活动窗口的窗口序号，新建一个窗口，然后激活前一个窗口。*/ function WinNum(){     let lwindowNum      lwindowNum = ActiveDocument.ActiveWindow.WindowNumber     NewWindow()     ActiveDocument.Windows.Item(lwindowNum).Activate() }  ` |

#### **Window.WindowState**

返回或设置指定文档窗口或任务窗口的状态。**WdWindowState** 类型，可读写。

**语法**

**express.WindowState**

*express*   一个代表 **Window** 对象的变量。

**说明**

**wdWindowStateNormal** 常量表示没有最大化或最小化的窗口。无法设置不活动窗口的状态。在设置窗口状态之前可用 **Activate** 方法激活窗口。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将没有最大化或最小化的活动窗口最大化。*/ function test() { if(ActiveDocument.ActiveWindow.WindowState == wdWindowStateNormal) {     ActiveDocument.ActiveWindow.WindowState = wdWindowStateMaximize } }` |

适用环境：web

适用平台：windows/linux