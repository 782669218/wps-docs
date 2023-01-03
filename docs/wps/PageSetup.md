#### **PageSetup**



代表页面设置说明。**PageSetup** 对象包含文档的所有页面设置属性（如左边距、下边距和纸张大小）。

**说明**

使用 **PageSetup** 属性可返回 **PageSetup** 对象。以下示例将活动文档的第一节设为横向并打印该文档。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.ActiveDocument.Sections.Item(1).PageSetup.Orientation = wdOrientLandscape     Application.ActiveDocument.PrintOut() }` |

以下示例设置命名为“Sales.doc”的文档的所有页边距。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let pages = Application.Documents.Item("Sales.doc").PageSetup     pages.LeftMargin = Application.InchesToPoints(0.75)     pages.RightMargin = Application.InchesToPoints(0.75)     pages.TopMargin = Application.InchesToPoints(1.5)     pages.BottomMargin = Application.InchesToPoints(1) }` |

**方法**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetAsTemplateDefault** | 将指定的页面设置格式设为活动文档和所有基于活动模板的新文档的默认格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TogglePortrait**       | 在文档或节的纵向和横向页面方向之间切换。                     |

**属性**

|                                                              | 名称                               | 说明                                                         |
| ------------------------------------------------------------ | ---------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                    | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BookFoldPrinting**               | 如果为 **True**，则 WPS 将文档打印为一系列小册子，这样可以折叠打印的页面并以书籍的形式阅读。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BookFoldPrintingSheets**         | 返回或设置一个 **Long** 类型的值，该值代表每一本小册子的页数。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BookFoldRevPrinting**            | 如果为 **True**，则 WPS 反转用于双向或亚洲语言文档的书籍折页打印的打印顺序。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BottomMargin**                   | 返回或设置页面底边与正文文本边界之间的距离（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CharsLine**                      | 返回或设置文档网格中每行的字符数。**Single** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                        | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DifferentFirstPageHeaderFooter** | 如果第一页使用了不同的页眉或页脚，则为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstPageTray**                  | 返回或设置用于文档或章节首页的纸盒。**WdPaperTray** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FooterDistance**                 | 返回或设置页脚与页面底边之间的距离（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Gutter**                         | 返回或设置文档或节的每一页的额外页边距的大小（以磅为单位），以备装订需要。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GutterPos**                      | 返回或设置文档中的装订线位于哪一侧。**WdGutterStyle** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GutterStyle**                    | 返回或设置 WPS 将装订线应用于当前文档时，基于从右向左的语言还是从左向右的语言。**WdGutterStyleOld** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeaderDistance**                 | 返回或设置页眉与页面顶端之间的距离（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LayoutMode**                     | 返回或设置当前文档的版式模式。**WdLayoutMode** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftMargin**                     | 返回或设置页面左边缘与正文左边界之间的距离（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineNumbering**                  | 返回或设置一个 **LineNumbering** 对象，该对象代表指定 **PageSetup** 对象的行号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LinesPage**                      | 返回或设置文档网格中每页的行数。**Single** 类型，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MirrorMargins**                  | 如果对开页的内外边距宽度相等，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OddAndEvenPagesHeaderFooter**    | 如果指定的 **PageSetup** 对象的奇数页和偶数页具有不同的页眉和页脚，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Orientation**                    | 返回或设置页面的方向。可读写 **WdOrientation** 类型。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OtherPagesTray**                 | 返回或设置用于文档或节中除第一页以外其他所有页的纸盒。**WdPaperTray** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageHeight**                     | 返回或设置页面高度（以磅为单位）。**Single** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageWidth**                      | 返回或设置页面宽度（以磅为单位）。**Single** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PaperSize**                      | 返回或设置纸张大小。**WdPaperSize** 类型，可读写。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                         | 返回一个 **Object** 类型值，该值代表指定 **PageSetup** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightMargin**                    | 返回或设置页面右边距与正文右边界之间的距离（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SectionDirection**               | 返回或设置指定节的阅读顺序和对齐方式。**WdSectionDirection** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SectionStart**                   | 返回或设置指定对象的分节符类型。**WdSectionStart** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowGrid**                       | 您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 **PageSetup** 对象的 **ShowGrid** 属性的信息，请查阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SuppressEndnotes**               | 如果在下一个没有隐藏尾注的节的末尾打印尾注，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextColumns**                    | 返回一个 **TextColumns** 集合，该集合代表指定 **PageSetup** 对象的文本栏集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TopMargin**                      | 返回或设置页面的上边缘与正文文本的上边界之间的距离（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TwoPagesOnOne**                  | 如果 WPS 将指定文档的每两页打印在同一张纸上，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalAlignment**              | 返回或设置文档或节中每页文本的垂直对齐方式。可读写 **WdVerticalAlignment** 类型。 |

**成员方法**

#### **PageSetup.SetAsTemplateDefault**

将指定的页面设置格式设为活动文档和所有基于活动模板的新文档的默认格式。

**语法**

**express.SetAsTemplateDefault()**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更改活动文档的左右页边距设置，然后将该页面设置格式设为默认格式。*/ function test() { let pagesetup = ActiveDocument.PageSetup pagesetup.LeftMargin = InchesToPoints(1) pagesetup.RightMargin = InchesToPoints(1) pagesetup.SetAsTemplateDefault() }` |

#### **PageSetup.TogglePortrait**

在文档或节的纵向和横向页面方向之间切换。

**语法**

**express.TogglePortrait()**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例改变活动文档的页面方向。*/ ActiveDocument.PageSetup.TogglePortrait()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将改变文档选定部分的所有节的页面方向。如果每节原来的方向与其他节不同，将会导致出错。*/ Selection.PageSetup.TogglePortrait()` |

**成员属性**

#### **PageSetup.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **PageSetup.BookFoldPrinting**

如果为 **True**，则 WPS 将文档打印为一系列小册子，这样可以折叠打印的页面并以书籍的形式阅读。**Boolean** 类型，可读写。

**语法**

**express.BookFoldPrinting**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档转换为以一页四开的形式打印的小册子。*/ function Booklet() {      let pages = ActiveDocument.PageSetup     pages.BookFoldPrinting = true     pages.BookFoldPrintingSheets = 4 }` |

#### **PageSetup.BookFoldPrintingSheets**

返回或设置一个 **Long** 类型的值，该值代表每一本小册子的页数。**Boolean** 类型，可读写。

**语法**

**express.BookFoldPrintingSheets**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档转换为以 16 页形式打印的小册子。*/ function Booklet() {      let pages = ActiveDocument.PageSetup     pages.BookFoldPrinting = true     pages.BookFoldPrintingSheets = 16 }` |

#### **PageSetup.BookFoldRevPrinting**

如果为 **True**，则 WPS 反转用于双向或亚洲语言文档的书籍折页打印的打印顺序。**Boolean** 类型，可读写。

**语法**

**express.BookFoldRevPrinting**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将从左向右的书籍打印切换为从右向左，并用于以一页 16 开打印的双向或亚洲语言文档。*/ function BookletRev() {      let pages = ActiveDocument.PageSetup     pages.BookFoldRevPrinting = true     pages.BookFoldPrintingSheets = 16 }` |

#### **PageSetup.BottomMargin**

返回或设置页面底边与正文文本边界之间的距离（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.BottomMargin**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的下边距设置为 72 磅（即 1 英寸），上边距设置为 2 英寸。InchesToPoints 方法用来将英寸转化为磅值。*/ function test() {   let pageSetup = Application.ActiveDocument.PageSetup   pageSetup.BottomMargin = 72   pageSetup.TopMargin = InchesToPoints(2) }  /*本示例将当前选定内容所有节的下边距设置为 2.5 英寸。*/ Application.Selection.PageSetup.BottomMargin = InchesToPoints(2.5)  /*本示例返回所选部分第一节的下边距。PointsToInches 方法用来将结果转换为英寸。*/ function test() {   let sngMargin = Application.Selection.Sections.Item(1).PageSetup.BottomMargin   alert(PointsToInches(sngMargin) + " inches") } ` |

#### **PageSetup.CharsLine**

返回或设置文档网格中每行的字符数。**Single** 类型，可读写。

**语法**

**express.CharsLine**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档每行字符数设置为 42。*/ ActiveDocument.PageSetup.CharsLine = 42` |

#### **PageSetup.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **PageSetup.DifferentFirstPageHeaderFooter**

如果第一页使用了不同的页眉或页脚，则为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.DifferentFirstPageHeaderFooter**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动文档的每一节，查找第一页是否有不同的页眉和页脚，如果有，则显示一条消息。*/ function test() { let sec = ActiveDocument.Sections for(let i=1;i<=sec.Count;i++) {     if(sec.Item(i).PageSetup.DifferentFirstPageHeaderFooter==true) {         MsgBox("Section" + i +"has different first page headers&footers.")     } } }` |

#### **PageSetup.FirstPageTray**

返回或设置用于文档或章节首页的纸盒。**WdPaperTray** 类型，可读写。

**语法**

**express.FirstPageTray**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置送纸盒，用来打印活动文档每节的首页。*/ ActiveDocument.PageSetup.FirstPageTray = wdPrinterLowerBin` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置送纸盒，用来打印所选部分每节的首页。*/ Selection.PageSetup.FirstPageTray = wdPrinterUpperBin` |

#### **PageSetup.FooterDistance**

返回或设置页脚与页面底边之间的距离（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.FooterDistance**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将页脚与页底边之间的距离设置为 0.5 英寸。InchesToPoints 方法用来将英寸转换为磅值。*/ ActiveDocument.PageSetup.FooterDistance = InchesToPoints(0.5)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定的所有节的页脚与页底边的距离设置为 1 英寸。*/ Selection.Range.PageSetup.FooterDistance = 72` |

#### **PageSetup.Gutter**

返回或设置文档或节的每一页的额外页边距的大小（以磅为单位），以备装订需要。**Single** 类型，可读写。

**语法**

**express.Gutter**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果将 **MirrorMargins** 属性设置为 **True**，则 **Gutter** 属性会将额外页边距添至内边距。否则，将额外页边距添至左边距。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的内边距增加 1 英寸（72 磅）。*/ function test() { ActiveDocument.PageSetup.MirrorMargins = true ActiveDocument.PageSetup.Gutter = 72 }` |

#### **PageSetup.GutterPos**

返回或设置文档中的装订线位于哪一侧。**WdGutterStyle** 类型，可读写。

**语法**

**express.GutterPos**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置装订线位于文档的右侧。*/ ActiveDocument.PageSetup.GutterPos = wdGutterPosRight` |

#### **PageSetup.GutterStyle**

返回或设置 WPS 将装订线应用于当前文档时，基于从右向左的语言还是从左向右的语言。**WdGutterStyleOld** 类型，可读写。

**语法**

**express.GutterStyle**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置当前文档使用从右向左语言文档的装订线样式。*/ ActiveDocument.PageSetup.GutterStyle = wdGutterStyleBidi` |

#### **PageSetup.HeaderDistance**

返回或设置页眉与页面顶端之间的距离（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.HeaderDistance**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示页眉与页面顶边之间的距离。PointsToInches 方法用来将磅值转换为英寸。*/ function test() { let sngDistance = ActiveDocument.PageSetup.HeaderDistance MsgBox(PointsToInches(sngDistance) + " inches") }` |

#### **PageSetup.LayoutMode**

返回或设置当前文档的版式模式。**WdLayoutMode** 类型，可读写。

**语法**

**express.LayoutMode**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档的版式模式，以便 WPS 可以自动将键入的文本与网格线对齐。*/ ActiveDocument.PageSetup.LayoutMode = wdLayoutModeGenko` |

#### **PageSetup.LeftMargin**

返回或设置页面左边缘与正文左边界之间的距离（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.LeftMargin**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果将**MirrorMargins**属性设置为 **True**，则 LeftMargin 属性将控制内边距的设置，而**MirrorMargins**属性将控制外边距的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档第二节的左边距设置为 1 英寸（72 磅）。*/ Application.ActiveDocument.Sections.Item(2).PageSetup.LeftMargin = 72` |

#### **PageSetup.LineNumbering**

返回或设置一个 **LineNumbering** 对象，该对象代表指定 **PageSetup** 对象的行号。

**语法**

**express.LineNumbering**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

只有在页面视图中才能看到行号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将为活动文档设置行号。*/ ActiveDocument.PageSetup.LineNumbering.Active = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为名为“MyDocument.doc”的文档设置行号。起始行号设置为 1，每 5 行显示一次行号，行号文档的各节中连续编号。*/ function test() { let myDoc = Documents.Item("MyDocument.doc") myDoc.PageSetup.LineNumbering.Active = true myDoc.PageSetup.LineNumbering.StartingNumber = 1 myDoc.PageSetup.LineNumbering.CountBy = 5 myDoc.PageSetup.LineNumbering.RestartMode = wdRestartContinuous }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的行号设置为与 MyDocument.doc 相同。*/ ActiveDocument.PageSetup.LineNumbering = Documents.Item("MyDocument.doc").PageSetup.LineNumbering    ` |

#### **PageSetup.LinesPage**

返回或设置文档网格中每页的行数。**Single** 类型，可读写。

**语法**

**express.LinesPage**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定活动文档的每页具有 15 行。*/ ActiveDocument.PageSetup.LinesPage = 15` |

#### **PageSetup.MirrorMargins**

如果对开页的内外边距宽度相等，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.MirrorMargins**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

**MirrorMargins** 属性可为 **True**、**False** 或 **wdUndefined**。如果将 **MirrorMargins** 属性设为 **True**，则 **LeftMargin** 属性将控制内边距的设置，而 **RightMargin** 属性将控制外边距的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的内边距设置为 1 英寸（72 磅），将外边距设置为 0.5 英寸。可以使用 InchesToPoints 方法将英寸转换为磅值。*/ function test() { let pagesetup = ActiveDocument.PageSetup pagesetup.MirrorMargins = true pagesetup.LeftMargin = 72 pagesetup.RightMargin = InchesToPoints(0.5) }` |

#### **PageSetup.OddAndEvenPagesHeaderFooter**

如果指定的 **PageSetup** 对象的奇数页和偶数页具有不同的页眉和页脚，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.OddAndEvenPagesHeaderFooter**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

**OddAndEvenPagesHeaderFooter** 属性值可以为 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Document1 的奇偶页创建不同的页眉和页脚。*/ function test() { let myDoc = Documents.Item("Document1") myDoc.PageSetup.OddAndEvenPagesHeaderFooter = true let primary = myDoc.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary) primary.Range.InsertAfter("Odd Header") let evenPages = myDoc.Sections.Item(1).Headers.Item(wdHeaderFooterEvenPages) evenPages.Range.InsertAfter("Even Header") }` |

#### **PageSetup.Orientation**

返回或设置页面的方向。可读写 **WdOrientation** 类型。

**语法**

**express.Orientation**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **WdOrientation** 常量可能不可用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例先改变文档“MyDocument.doc”的方向，再打印该文档。然后本示例将文档的方向恢复为纵向。*/ function test() { let myDoc = Documents.Item("MyDocument.doc") myDoc.PageSetup.Orientation = wdOrientLandscape myDoc.PrintOut() myDoc.PageSetup.Orientation = wdOrientPortrait }  ` |

#### **PageSetup.OtherPagesTray**

返回或设置用于文档或节中除第一页以外其他所有页的纸盒。**WdPaperTray** 类型，可读写。

**语法**

**express.OtherPagesTray**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置纸盒用于打印活动文档各节除第一页以外的所有页。*/ ActiveDocument.PageSetup.OtherPagesTray = wdPrinterUpperBin` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置纸盒用于打印选定内容各节中除第一页以外的所有页。*/ Selection.PageSetup.OtherPagesTray = wdPrinterLowerBin` |

#### **PageSetup.PageHeight**

返回或设置页面高度（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.PageHeight**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果设置了 **PageHeight** 属性，则 **PaperSize** 属性会改为 [wdPaperCustom](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdPaperSize%20%E6%9E%9A%E4%B8%BE.html)。使用 **PaperSize** 属性可设置预定义大小的纸张的页面高度和宽度，例如 Letter 或 A4 纸。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的页面高度设置为 9 英寸。*/ function test() {   Application.ActiveDocument.PageSetup.PageHeight = InchesToPoints(9)   Application.ActiveDocument.PageSetup.PageWidth = InchesToPoints(7) }` |

#### **PageSetup.PageWidth**

返回或设置页面宽度（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.PageWidth**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果设置了 **PageWidth** 属性，则 **PaperSize** 属性将改为 [wdPaperCustom](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdPaperSize%20%E6%9E%9A%E4%B8%BE.html)。使用 **PaperSize** 属性可设置预定义大小的纸张的页面高度和宽度，例如 Letter 或 A4 纸。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回 Document1 的页面宽度。可用 PointsToInches 方法将磅值转换为英寸。*/ function test() {   let doc1set = Application.Documents.Item("Document1.docx").PageSetup   alert("The page width is "+PointsToInches(doc1set.PageWidth)+" inches.") } ` |

#### **PageSetup.PaperSize**

返回或设置纸张大小。**WdPaperSize** 类型，可读写。

**语法**

**express.PaperSize**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

设置 **PageHeight** 或 **PageWidth** 属性会将 **PaperSize** 属性更改为 **wdPaperCustom**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将第一个文档的纸张大小设置为 legal 型。*/ Application.Documents.Item(1).PageSetup.PaperSize = wdPaperLegal` |

#### **PageSetup.Parent**

返回一个 **Object** 类型值，该值代表指定 **PageSetup** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.RightMargin**

返回或设置页面右边距与正文右边界之间的距离（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.RightMargin**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

如果将 **MirrorMargins** 属性设置为 **True**，则 **RightMargin** 属性将控制外边距的设置，而 **LeftMargin** 属性控制内边距的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档的右边距设置。可以使用 PointsToInches 方法将结果转换为英寸。*/ alert("The right margin is set to " + PointsToInches(Application.ActiveDocument.PageSetup.RightMargin) + " inches.")  /*本示例设置选定区域第二节的右边距。可以使用 InchesToPoints 方法将英寸转换为磅值。*/ Application.Selection.Sections.Item(2).PageSetup.RightMargin = InchesToPoints(1)` |

#### **PageSetup.SectionDirection**

返回或设置指定节的阅读顺序和对齐方式。**WdSectionDirection** 类型，可读写。

**语法**

**express.SectionDirection**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的第一节的方向设置为从右向左。*/ ActiveDocument.Sections.Item(1).PageSetup.SectionDirection = wdSectionDirectionRtl` |

#### **PageSetup.SectionStart**

返回或设置指定对象的分节符类型。**WdSectionStart** 类型，可读写。

**语法**

**express.SectionStart**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中所有节的分节符都改为连续的。*/ ActiveDocument.PageSetup.SectionStart = wdSectionContinuous` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回 MyDoc.doc 中第二节开始部分使用的分节符类型，并将其应用于活动文档中的所有节。*/ function test() { mytype = Documents.Item("MyDoc.doc").Sections.Item(2).PageSetup.SectionStart ActiveDocument.PageSetup.SectionStart = mytype }  ` |

#### **PageSetup.ShowGrid**

您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 **PageSetup** 对象的 **ShowGrid** 属性的信息，请查阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

**语法**

**express.ShowGrid**

*express*   一个代表 **PageSetup** 对象的变量。

#### **PageSetup.SuppressEndnotes**

如果在下一个没有隐藏尾注的节的末尾打印尾注，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.SuppressEndnotes**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

在该节的尾注之前打印隐藏的尾注。仅当 **Location** 属性设为 **wdEndOfSection** 时，该属性才有效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏活动文档中第一节的尾注。*/ function test() { if(ActiveDocument.Endnotes.Location == wdEndOfSection) {     ActiveDocument.Sections.Item(1).PageSetup.SuppressEndnotes = true } }` |

#### **PageSetup.TextColumns**

返回一个 **TextColumns** 集合，该集合代表指定 **PageSetup** 对象的文本栏集合。

**语法**

**express.TextColumns**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

集合中应始终至少包含一个文本列。在创建新的文本列时，会向集合添加一列。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档中的第二节创建四个等宽的文本栏。*/ function test() { let textcolumns = ActiveDocument.Sections.Item(2).PageSetup.TextColumns textcolumns.SetCount(3) textcolumns.Add(null,null,true) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个具有两个文本栏的文档。第一个文本栏宽 1.5 英寸，第二个文本栏宽 3 英寸。*/ function test() { let myDoc = Documents.Add() let textcolumns = myDoc.PageSetup.TextColumns textcolumns.SetCount(1) textcolumns.Add(InchesToPoints(3)) let item = myDoc.PageSetup.TextColumns.Item(1) item.Width = InchesToPoints(1.5) item.SpaceAfter = InchesToPoints(0.5) } ` |

#### **PageSetup.TopMargin**

返回或设置页面的上边缘与正文文本的上边界之间的距离（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.TopMargin**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一节的上边距设置为 72 磅（1 英寸）。*/ Application.ActiveDocument.Sections.Item(1).PageSetup.TopMargin = 72` |

#### **PageSetup.TwoPagesOnOne**

如果 WPS 将指定文档的每两页打印在同一张纸上，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.TwoPagesOnOne**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为将活动文档的每两页打印在同一张纸上。*/ ActiveDocument.PageSetup.TwoPagesOnOne = true` |

#### **PageSetup.VerticalAlignment**

返回或设置文档或节中每页文本的垂直对齐方式。可读写 **WdVerticalAlignment** 类型。

**语法**

**express.VerticalAlignment**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例改变第一个文档的垂直对齐方式，使文本在上边距和下边距之间居中。*/ Documents.Item(1).PageSetup.VerticalAlignment = wdAlignVerticalCenter` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例新建一个文档，并将同一段落插入 10 次。然后设置新文档的垂直对齐方式，使 10 个段落在上边距和下边距之间等距排列（两端对齐）。*/ function test() { let myDoc = Documents.Add() for(let i= 1;i<=9;i++) {     myDoc.Content.InsertAfter("This is a sentence.")     myDoc.Content.InsertParagraphAfter() } myDoc.Content.InsertAfter("This is a sentence.") myDoc.PageSetup.VerticalAlignment = wdAlignVerticalJustify }` |

适用环境：web

适用平台：windows/linux