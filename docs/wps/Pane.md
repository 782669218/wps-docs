#### **Pane**



代表一个窗格。**Pane** 对象是 **Panes** 集合的一个成员。 **Panes** 集合包含单个窗口的所有窗格。

**说明**

使用 **Panes**(*Index*) 可返回一个 **Pane** 对象，其中 *Index* 为索引号。以下示例关闭活动窗格。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { if(ActiveDocument.ActiveWindow.Panes.Count >= 2) {     ActiveDocument.ActiveWindow.ActivePane.Close() } }` |

使用 **Add** 方法或 **Split** 属性可添加一个窗格。以下示例按当前窗口大小的 20% 拆分活动窗口。

| 示例代码复制                                |
| ------------------------------------------- |
| `ActiveDocument.ActiveWindow.Panes.Add(20)` |

以下示例将活动窗口一分为二。

| 示例代码复制                               |
| ------------------------------------------ |
| `ActiveDocument.ActiveWindow.Split = true` |

您可以使用 **SplitSpecial** 属性显示单独窗格中的批注、脚注或尾注。

当窗口被拆分或在非页面视图下显示脚注、批注等信息时，窗口具有多个窗格。以下示例在普通视图中显示批注窗格，然后提示关闭该窗格。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/**/ function test() { ActiveDocument.ActiveWindow.View.Type = wdNormalView if(ActiveDocument.Comments.Count >= 1) {     ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneComments     let response = MsgBox("Do you want to close the comments pane?",jsYesNo)     if(response == jsResultYes) {         ActiveDocument.ActiveWindow.ActivePane.Close()     } } }` |

**方法**

|                                                              | 名称              | 说明                                                       |
| ------------------------------------------------------------ | ----------------- | ---------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**      | 激活指定的窗格。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutoScroll**    | 在指定的窗格中自动滚动。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Close**         | 关闭指定的邮件合并数据源、窗格或任务。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LargeScroll**   | 将窗口或窗格滚动指定的屏数。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NewFrameset**   | 基于指定的窗格新建一个框架页。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PageScroll**    | 将指定的窗口按页滚动。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SmallScroll**   | 将窗口滚动指定的行数。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TOCInFrameset** | 基于指定文档创建一个目录，并将其放入框架页左边的新框架中。 |

**属性**

|                                                              | 名称                          | 说明                                                         |
| ------------------------------------------------------------ | ----------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**               | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BrowseWidth**               | 返回指定窗格中正文换行区域的宽度（以磅为单位）。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                   | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayRulers**             | 如果在指定窗格中显示标尺，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayVerticalRuler**      | 如果在指定窗格中显示垂直标尺，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Document**                  | 返回一个 **Document** 对象，该对象与指定窗格相关。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Frameset**                  | 返回一个 **Frameset** 对象，该对象代表一个整体框架页或框架页的单个框架。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalPercentScrolled** | 返回或设置水平滚动条位置（以文档宽度的百分比表示）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**                     | 返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MinimumFontSize**           | 返回或设置在指定窗格中显示的最小字号（以磅为单位）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Next**                      | 返回一个 **Pane** 对象，该对象代表集合中的下一个文档窗格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Pages**                     | 返回一个 **Pages** 集合，该集合代表文档中的页面。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                    | 返回一个 **Object** 类型值，该值代表指定 **Pane** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Previous**                  | 返回一个 **Pane** 对象，该对象代表集合中的上一个文档窗格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Selection**                 | 返回 **Selection** 对象，该对象代表文档窗格中的所选内容或插入点。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalPercentScrolled**   | 返回或设置垂直滚动条的位置（以文档长度的百分比表示）。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **View**                      | 返回一个 **View** 对象，该对象代表指定窗格的视图。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Zooms**                     | 返回一个 **Zooms** 集合，该集合代表每个视图（如普通视图、大纲视图和页面视图）的缩放选项。 |

**成员方法**

#### **Pane.Activate**

激活指定的窗格。

**语法**

**express.Activate()**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例拆分活动窗口，然后激活第一个窗格。*/ function SplitWindow() {     let activewindow = ActiveDocument.ActiveWindow     activewindow.SplitVertical = 50     activewindow.Panes.Item(1).Activate() }` |

#### **Pane.AutoScroll**

在指定的窗格中自动滚动。

**语法**

**express.AutoScroll(Velocity)**

*express*   一个代表 **Pane** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Velocity* | 必选          | **Long**     | 滚动的速度。取值范围从 -100 到 100。使用 -100 表示以最快速度向后滚动，使用 100 表示以最快速度向前滚动。 |

**说明**

该方法将连续执行，直到通过按键或单击鼠标的方式手动将其停止。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使文档在活动窗格中慢速向后滚动。*/ ActiveDocument.ActiveWindow.ActivePane.AutoScroll(-20)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使文档在活动窗格中以最快速度向前滚动。*/ ActiveDocument.ActiveWindow.ActivePane.AutoScroll(100)` |

#### **Pane.Close**

关闭指定的邮件合并数据源、窗格或任务。

**语法**

**express.Close()**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在拆分活动窗口时关闭活动窗格。*/ function test() { if(ActiveDocument.ActiveWindow.Panes.Count >= 2) {     ActiveDocument.ActiveWindow.ActivePane.Close() } }` |

#### **Pane.LargeScroll**

将窗口或窗格滚动指定的屏数。

**语法**

**express.LargeScroll(Down, Up, ToRight, ToLeft)**

*express*   一个代表 **Pane** 对象的变量。

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

#### **Pane.NewFrameset**

基于指定的窗格新建一个框架页。

**语法**

**express.NewFrameset()**

*express*   一个代表 **Pane** 对象的变量。

**说明**

有关创建框架页的详细信息，请参阅 创建框架页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将打开一个名为“Temp.doc”的文件，然后创建一个新的框架页，该框架页中唯一的一个框架中包含“Temp.doc”。*/ function test() { Documents.Open("C:\\Documents\\Temp.doc") ActiveDocument.ActiveWindow.ActivePane.NewFrameset() }  ` |

#### **Pane.PageScroll**

将指定的窗口按页滚动。

**语法**

**express.PageScroll(Down, Up)**

*express*   一个代表 **Pane** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                           |
| -------- | ------------- | ------------ | -------------------------------------------------- |
| *Down*   | 可选          | **Variant**  | 向下滚动的页数。如果省略此参数，则参数值设置为 1。 |
| *Up*     | 可选          | **Variant**  | Variant 向上滚动的页数。                           |

**说明**

**PageScroll** 方法仅用于页面视图或 Web 版式视图。此方法不影响插入点的位置。

如果指定了 *Down* 和 *Up* 两个参数，则窗口滚动页数为这两个参数的差。例如，如果 *Down* 为 2，*Up* 为 4，则窗口向上滚动两页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例使活动窗格向上滚动一页。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdPrintView ActiveDocument.ActiveWindow.ActivePane.PageScroll(null,1) }` |

#### **Pane.SmallScroll**

将窗口滚动指定的行数。

**语法**

**express.SmallScroll(Down, Up, ToRight, ToLeft)**

*express*   一个代表 **Pane** 对象的变量。

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

#### **Pane.TOCInFrameset**

基于指定文档创建一个目录，并将其放入框架页左边的新框架中。

**语法**

**express.TOCInFrameset()**

*express*   一个代表 **Pane** 对象的变量。

**说明**

有关创建框架页的详细信息，请参阅 创建框架页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：打开一个名为“Proposal.doc”的文件，基于该文件创建一个框架页，并添加一个含有该文件目录的框架（位于页面的左侧）。*/ function test() { Documents.Open("C:\\Documents\\Proposal.doc") ActiveDocument.ActiveWindow.ActivePane.NewFrameset() ActiveDocument.ActiveWindow.ActivePane.TOCInFrameset() }` |

**成员属性**

#### **Pane.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Pane** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Pane.BrowseWidth**

返回指定窗格中正文换行区域的宽度（以磅为单位）。**Long** 类型，只读。

**语法**

**express.BrowseWidth**

*express*   一个代表 **Pane** 对象的变量。

**说明**

此属性仅用于 Web 版式视图。

#### **Pane.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Pane** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Pane.DisplayRulers**

如果在指定窗格中显示标尺，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayRulers**

*express*   一个代表 **Pane** 对象的变量。

**说明**

**DisplayRulers** 属性相当于**“视图”**菜单上的**“标尺”**命令。如果 **DisplayRulers** 为 **False**，则无论 **DisplayVerticalRuler** 属性处于何状态，都不显示水平和垂直标尺。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗格切换到页面视图并显示水平和垂直标尺。*/ function test() { let activepane = ActiveDocument.ActiveWindow.ActivePane activepane.View.Type = wdPrintView activepane.DisplayRulers = true activepane.DisplayVerticalRuler = true }` |

#### **Pane.DisplayVerticalRuler**

如果在指定窗格中显示垂直标尺，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayVerticalRuler**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗格切换到页面视图并显示水平和垂直标尺。*/ function test() { let activepane = ActiveDocument.ActiveWindow.ActivePane activepane.View.Type = wdPrintView activepane.DisplayRulers = true activepane.DisplayVerticalRuler = true }  ` |

#### **Pane.Document**

返回一个 **Document** 对象，该对象与指定窗格相关。只读。

**语法**

**express.Document**

*express*   一个代表 **Pane** 对象的变量。

#### **Pane.Frameset**

返回一个 **Frameset** 对象，该对象代表一个整体框架页或框架页的单个框架。只读。

**语法**

**express.Frameset**

*express*   一个代表 **Pane** 对象的变量。

**说明**

有关创建框架页的详细信息，请参阅 创建框架页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例紧邻指定框架的右侧添加一个新框架。*/ ActiveDocument.ActiveWindow.ActivePane.Frameset.AddNewFrame(wdFramesetNewRight)    ` |

#### **Pane.HorizontalPercentScrolled**

返回或设置水平滚动条位置（以文档宽度的百分比表示）。**Long** 类型，可读写。

**语法**

**express.HorizontalPercentScrolled**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 Document1 窗口的活动窗格水平滚动至最左边。*/ function test() { Windows.Item("Document1").Activate() Windows.Item("Document1").ActivePane.HorizontalPercentScrolled = 0 }` |

#### **Pane.Index**

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

*express*   一个代表 **Pane** 对象的变量。

#### **Pane.MinimumFontSize**

返回或设置在指定窗格中显示的最小字号（以磅为单位）。**Long** 类型，可读写。

**语法**

**express.MinimumFontSize**

*express*   一个代表 **Pane** 对象的变量。

**说明**

该属性将只影响 Web 版式视图中显示的文本，而**“格式”**工具栏上显示的磅值和用于打印的磅值则不会改变。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口设置为大纲视图，并将活动窗格中的最小字号设置为 12 磅。*/ function test() { let activewindow = ActiveDocument.ActiveWindow activewindow.View.Type = wdWebView activewindow.ActivePane.MinimumFontSize = 12 }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回活动窗格的最小字号。*/ MsgBox(ActiveDocument.ActiveWindow.ActivePane.MinimumFontSize)` |

#### **Pane.Next**

返回一个 **Pane** 对象，该对象代表集合中的下一个文档窗格。只读。

**语法**

**express.Next**

*express*   一个代表 **Pane** 对象的变量。

#### **Pane.Pages**

返回一个 **Pages** 集合，该集合代表文档中的页面。

**语法**

**express.Pages**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例从活动文档左上角 0.5 英寸处穿过页面向页面右下角（距页面右边缘和下边缘 0.5 英寸处）创建一条线。*/ function test() { let objPage = ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1)  //Add new line to document ActiveDocument.Shapes.AddLine(InchesToPoints(0.5), InchesToPoints(0.5), objPage.Width - InchesToPoints(0.5), objPage.Height - InchesToPoints(0.5))     }` |

#### **Pane.Parent**

返回一个 **Object** 类型值，该值代表指定 **Pane** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Pane** 对象的变量。

#### **Pane.Previous**

返回一个 **Pane** 对象，该对象代表集合中的上一个文档窗格。只读。

**语法**

**express.Previous**

*express*   一个代表 **Pane** 对象的变量。

#### **Pane.Selection**

返回 **Selection** 对象，该对象代表文档窗格中的所选内容或插入点。只读。

**语法**

**express.Selection**

*express*   一个代表 **Pane** 对象的变量。

#### **Pane.VerticalPercentScrolled**

返回或设置垂直滚动条的位置（以文档长度的百分比表示）。可读写 **Long** 类型。

**语法**

**express.VerticalPercentScrolled**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 Document1 窗口的活动窗格垂直滚动至结尾。*/ function test() { let document = Windows.Item("Document1") document.Activate() document.ActivePane.VerticalPercentScrolled = 100 }` |

#### **Pane.View**

返回一个 **View** 对象，该对象代表指定窗格的视图。

**语法**

**express.View**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示与 Windows 集合中的第一个窗口相关联的窗格的所有非打印字符。*/ function test() { for(let i = 1; i <= Windows.Item(1).Panes.Count; i++) {     Windows.Item(1).Panes.Item(i).View.ShowAll = true } }` |

#### **Pane.Zooms**

返回一个 **Zooms** 集合，该集合代表每个视图（如普通视图、大纲视图和页面视图）的缩放选项。

**语法**

**express.Zooms**

*express*   一个代表 **Pane** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将在页面视图中打开的所有窗口的显示比例设为 100%。*/ function test() { for(let i = 1; i <= Windows.Count; i++) {     Windows.Item(i).ActivePane.Zooms.Item(wdPrintView).Percentage = 100 } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置页面视图的显示比例，使整个页面可见。*/ ActiveDocument.ActiveWindow.Panes.Item(1).Zooms.Item(wdPrintView).PageFit = wdPageFitFullPage    ` |

适用环境：web

适用平台：windows/linux