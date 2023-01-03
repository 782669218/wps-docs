#### **View**



包含窗口或窗格的视图属性（如全部显示、域底纹和网格线）。

**说明**

使用 **View** 属性可返回 **View** 对象。以下示例设置活动窗口的视图选项。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let view = Application.ActiveDocument.ActiveWindow.View view.ShowAll = true view.TableGridlines = true view.WrapToWindow = false` |

使用 

Type

 属性可更改视图。以下示例将活动窗口切换至普通视图。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.ActiveWindow.View.Type = wdNormalView` |

使用 

Percentage

 属性可更改屏幕上文本的尺寸。以下示例将屏幕上文本的显示比例放大到 120%。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.ActiveWindow.View.Zoom.Percentage = 120` |

使用 

SeekView

 属性可查看批注、尾注、脚注或者文档页眉或页脚。以下示例在页面视图中显示活动窗口中的当前页脚。示例代码复制`Application.ActiveDocument.ActiveWindow.View.Type = wdPrintView  Application.ActiveDocument.ActiveWindow.View.SeekView = wdSeekCurrentPageFooter`

**方法**

|                                                              | 名称                     | 说明                                                   |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CollapseOutline**      | 将选定内容或指定范围内的正文折叠一个标题级别。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ExpandOutline**        | 按照一个标题级别展开选定内容下的文本。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PreviousHeaderFooter** | 根据视图中显示的是页眉还是页脚，移动到前一页眉或页脚。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowAllHeadings**      | 在显示所有文本（包括标题和正文）和仅显示标题之间切换。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowHeading**          | 显示所有高于指定标题级别的标题，隐藏子标题和正文。     |

**属性**

|                                                              | 名称                                    | 说明                                                         |
| ------------------------------------------------------------ | --------------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                         | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ConflictMode**                        | 如果文档位于冲突模式视图中，则该属性值为 **True**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                             | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayBackgrounds**                  | 返回或设置一个 **Boolean** 类型的值，该值代表在页面视图中显示文档时，是否显示背景色和图像。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayPageBoundaries**               | 如果为 **True**，则显示文档顶部和底部页边距（空白区域）并在页面间显示灰色区域（灰色间隔）。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplaySmartTags**                    | 如果为 **True**，则 WPS 在文档中的智能标记下方显示一条下划线。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Draft**                               | 如果窗口中的所有文本以具有最少格式信息的相同的 sans-serif 字体显示以加速显示，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FieldShading**                        | 返回或设置域的屏幕底纹。**WdFieldShading** 类型，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FullScreen**                          | 如果窗口为全屏显示，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Magnifier**                           | 如果为 **True**，则在打印预览中将鼠标指针显示为放大镜形状，表示用户可以通过单击来放大页面的某个区域或通过缩小显示比例查看整页或展开多页。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MailMergeDataView**                   | 如果在指定窗口中显示邮件合并数据，而不显示邮件合并域，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkupMode**                          | 返回或设置一个 **WdRevisionsMode** 常量，该常量代表修订的显示模式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Panning**                             | 返回或设置 **Boolean** 类型的值，该值表示 WPS 是否处于全景模式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                              | 返回一个 **Object** 类型值，该值代表指定 **View** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingLayout**                       | 设置或返回一个 **Boolean** 类型的值，该值表示是否在阅读版式视图中查看文档。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingLayoutActualView**             | 设置或返回一个 **Boolean** 类型的值，该值代表是否使用与打印页面相同的版式在阅读版式视图中显示页面。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingLayoutAllowEditing**           | 返回 **Boolean** 值，该值表示是否允许在阅读布局模式下编辑文本。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingLayoutAllowMultiplePages**     | 设置或返回一个 **Boolean** 类型的值，该值代表阅读版式视图是否并排显示两页。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingLayoutTruncateMargins**        | 返回或设置一个 **WdReadingLayoutMargin** 常量，该常量代表当在阅读版式中查看文档时是显示还是隐藏边距。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reviewers**                           | 返回一个 **Reviewers** 对象，该对象代表所有审阅者。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisionsBalloonShowConnectingLines** | 如果为 **True**，则 WPS 显示从文本到修订和批注气球之间的连接线。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisionsBalloonSide**                | 设置或返回一个 **WdRevisionsBalloonMargin** 常量，指定 WPS 是否在文档的左边距或右边距中显示修订气球。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisionsBalloonWidth**               | 设置或返回一个 **Single** 类型的值，该值代表 WPS 中用于指定修订气球宽度的全局设置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisionsBalloonWidthType**           | 设置或返回一个 **WdRevisionsBalloonWidthType** 常量，该常量代表指定 WPS 如何测量修订气球的宽度的全局设置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisionsView**                       | 设置或返回一个代表全局选项的 **WdRevisionsView** 常量，该选项用于指定 WPS 是显示文档的原始版本还是应用了修订和格式更改的版本。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SeekView**                            | 返回或设置在页面视图中显示的文档元素。**WdSeekView** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShadeEditableRanges**                 | 返回或设置一个 **Long** 类型的值，该值代表是否将底纹应用于文档中用户有权修改的区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAll**                             | 如果显示所有非打印字符（如隐藏文字、制表符、空格和段落标记），则该属性值为 **True**。可读写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAnimation**                       | 如果显示文本动画，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowBookmarks**                       | 如果在每个书签的首尾显示方括号，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowComments**                        | 如果为 **True**，则 WPS 在文档中显示批注。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowCropMarks**                       | 返回或设置 **Boolean** 值，该值表示是否在页面的角显示裁剪标记，以指示页边距的位置。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowDrawings**                        | 如果在页面视图中显示由绘图工具创建的对象，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowFieldCodes**                      | 如果显示域代码，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowFirstLineOnly**                   | 如果在大纲视图中仅显示正文每段的第一行，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowFormat**                          | 如果在大纲视图中显示字符格式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowFormatChanges**                   | 如果为 **True**，则在启用“修订”功能后，WPS 显示对文档所作的格式更改。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowHiddenText**                      | 如果显示具有隐藏格式的文本，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowHighlight**                       | 如果显示和打印文档中突出显示格式的文本，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowHyphens**                         | 如果为 **True**，则显示可选连字符，该连字符指明在何处断开处于行尾的字。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowInkAnnotations**                  | 返回或设置一个 **Boolean** 类型的值，该值表示显示还是隐藏手写墨迹注释。如果值为 **True**，则显示墨迹注释。如果值为 **False**，则隐藏墨迹注释。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowInsertionsAndDeletions**          | 如果为 **True**，则 WPS 显示用“修订”功能在文档中插入和删除的文本。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowMainTextLayer**                   | 如果在显示页眉和页脚区域时同时显示指定文档的文字，则该属性值为 **True**。该属性等效于**“页眉和页脚”**工具栏上的**“显示/隐藏文档文字”**按钮。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowMarkupAreaHighlight**             | 返回或设置 **Boolean** 值，该值表示显示修订和注释气球的标记区域是否带底纹。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowObjectAnchors**                   | 如果在可在页面视图中定位的项目附近显示对象锁定标记，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowOptionalBreaks**                  | 如果 WPS 显示可选换行符，则该属性值为 **True**。**Boolean** 类型，可读写 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowOtherAuthors**                    | 如果应在文档中显示其他作者，则该属性值为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowParagraphs**                      | 如果显示段落标记，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowPicturePlaceHolders**             | 如果将空白方框作为图片的占位符进行显示，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowRevisionsAndComments**            | 如果为 **True**，则 WPS 显示使用“修订”功能对文档所作的修订和批注。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowSpaces**                          | 如果显示空格字符，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowTabs**                            | 如果显示制表符，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowTextBoundaries**                  | 如果在页面视图中显示页边距、文本列、对象和框架周围的虚线，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowXMLMarkup**                       | 返回一个 **Long** 类型的值，该值表示是否在文档中显示 XML 标记。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SplitSpecial**                        | 返回或设置活动窗口的窗格。**WdSpecialPane** 类型，可读写。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TableGridlines**                      | 如果显示表格虚框，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                                | 返回或设置视图的类型。可读/写 [WdViewType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WrapToWindow**                        | 如果文字在文档窗口的右边缘换行，而不是在右边距或文本栏的右边界，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Zoom**                                | 返回一个 [Zoom ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Zoom/Zoom%20.htm#jsObject_Zoom)对象，该对象代表指定视图的显示比例。 |

**成员方法**

#### **View.CollapseOutline**

将选定内容或指定范围内的正文折叠一个标题级别。

**语法**

**express.CollapseOutline(Range)**

*express*   一个代表 **View** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**     | **说明**                                                 |
| -------- | ------------- | ---------------- | -------------------------------------------------------- |
| *Range*  | 可选          | **Range object** | 需要折叠的段落范围。如果省略此参数，则折叠全部选定内容。 |

**说明**

如果文档未处于大纲或主控文档视图，将导致出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的第二段设置“标题 2”样式，将活动窗口切换至大纲视图，并折叠第二段的正文。*/ function test() { ActiveDocument.Paragraphs.Item(2).Style = wdStyleHeading2  ActiveDocument.ActiveWindow.View.Type = wdOutlineView ActiveDocument.ActiveWindow.View.CollapseOutline(ActiveDocument.Paragraphs.Item(2).Range) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将文档中每个标题折叠一个级别。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdOutlineView ActiveDocument.ActiveWindow.View.CollapseOutline(ActiveDocument.Content) }` |

#### **View.ExpandOutline**

按照一个标题级别展开选定内容下的文本。

**语法**

**express.ExpandOutline(Range)**

*express*   一个代表 **View** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                 |
| -------- | ------------- | ------------ | -------------------------------------------------------- |
| *Range*  | 可选          | **Range**    | 要展开的段落的范围。如果省略此参数，则展开全部选定内容。 |

**说明**

如果文档未处于大纲或主控文档视图，将导致出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将文档中每个标题展开一个级别。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdOutlineView ActiveDocument.ActiveWindow.View.ExpandOutline(ActiveDocument.Content) }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例展开 Document2 窗口中的活动段落。*/ function test() { let myDoc = Windows.Item("Document2")  myDoc.Activate() myDoc.View.Type = wdOutlineView myDoc.View.ExpandOutline() }  ` |

#### **View.PreviousHeaderFooter**

根据视图中显示的是页眉还是页脚，移动到前一页眉或页脚。

**语法**

**express.PreviousHeaderFooter()**

*express*   一个代表 **View** 对象的变量。

**说明**

如果视图显示的是页眉，则该方法移至当前节的前一个页眉（例如，从偶数页眉到奇数页眉）或者前一节的最后一个页眉。如果视图显示页脚，则移至前一个页脚。

| 注释                                                         |
| ------------------------------------------------------------ |
| 如果视图显示文档的第一节的第一个页眉或页脚，或者如果根本不显示页眉或页脚，则导致出错。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例插入一个偶数页分节符，将活动窗口切换到页面视图，显示当前页眉，然后切换到前一个页眉。*/ function test() { Selection.Collapse(wdCollapseStart) Selection.InsertBreak(wdSectionBreakEvenPage)  ActiveDocument.ActiveWindow.View.Type = wdPrintView ActiveDocument.ActiveWindow.View.SeekView = wdSeekCurrentPageHeader ActiveDocument.ActiveWindow.View.PreviousHeaderFooter() }` |

#### **View.ShowAllHeadings**

在显示所有文本（包括标题和正文）和仅显示标题之间切换。

**语法**

**express.ShowAllHeadings()**

*express*   一个代表 **View** 对象的变量。

**说明**

如果不是大纲视图或者主控文档视图，此方法将导致出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在大纲视图中使用 ShowHeading 方法显示所有标题（不显示正文），然后切换显示所有文本（包括标题和正文）。*/ function test() { let view = ActiveDocument.ActiveWindow.View  view.Type = wdOutlineView view.ShowHeading(9) view.ShowAllHeadings() }　　　　` |

#### **View.ShowHeading**

显示所有高于指定标题级别的标题，隐藏子标题和正文。

**语法**

**express.ShowHeading(Level)**

*express*   一个代表 **View** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                   |
| -------- | ------------- | ------------ | ------------------------------------------ |
| *Level*  | 必选          | **Long**     | 指大纲标题的级别（可取 1 到 9 之间数值）。 |

**说明**

如果不是大纲视图或者主控文档视图，此方法将导致出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口切换到大纲视图，并显示所有以“标题 1”样式为格式的文本。隐藏正文和其他类型的标题。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdOutlineView ActiveDocument.ActiveWindow.View.ShowHeading(1) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Document1 窗口切换到大纲视图，并显示所有以“标题 1”、“标题 2”或“标题 3”样式为格式的文本。*/ function test() { let myDoc = Windows.Item("Document1").View myDoc.Type = wdOutlineView myDoc.ShowHeading(3) }` |

**成员属性**

#### **View.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **View** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **View.ConflictMode**

如果文档位于冲突模式视图中，则该属性值为 **True**。可读/写。

**语法**

**express.ConflictMode**

*express*   一个代表 **View** 对象的变量。

#### **View.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **View** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **View.DisplayBackgrounds**

返回或设置一个 **Boolean** 类型的值，该值代表在页面视图中显示文档时，是否显示背景色和图像。

**语法**

**express.DisplayBackgrounds**

*express*   一个代表 **View** 对象的变量。

**说明**

对应于**“选项”**对话框的**“视图”**选项卡上的**“背景色和图像(仅页面视图)”**选项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在页面视图中显示活动文档时，隐藏背景色和图像。*/ ActiveDocument.ActiveWindow.View.DisplayBackgrounds = false` |

#### **View.DisplayPageBoundaries**

如果为 **True**，则显示文档顶部和底部页边距（空白区域）并在页面间显示灰色区域（灰色间隔）。**Boolean** 类型，可读写。

**语法**

**express.DisplayPageBoundaries**

*express*   一个代表 **View** 对象的变量。

**说明**

如果为 **False**，则隐藏这些白色和灰色区域，使页面形成很长的一页。默认值为 **True**。

该功能仅在页面视图中可用，并仅影响页面顶部和底部的灰色间隔，而不影响页面的左右两边。该设置影响指定窗口中的文档。保存该文档时，同时保存该设置的状态。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前视图更改为“页面视图”并隐藏文档页面之间的白色和灰色间隔。*/ function WhiteSpace(){     ActiveWindow.View.Type = wdPrintView     ActiveWindow.View.DisplayPageBoundaries = false }` |

#### **View.DisplaySmartTags**

如果为 **True**，则 WPS 在文档中的智能标记下方显示一条下划线。**Boolean** 类型，可读写。

**语法**

**express.DisplaySmartTags**

*express*   一个代表 **View** 对象的变量。

**说明**

在文档中用下划线标记智能标记。将 **DisplaySmartTags** 属性设为 **False** 不能删除智能标记；该操作仅关闭下划线的显示。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动视图中关闭智能标记下方下划线的显示。*/ function DontShowSmartTags(){     ActiveWindow.View.DisplaySmartTags = false }` |

#### **View.Draft**

如果窗口中的所有文本以具有最少格式信息的相同的 sans-serif 字体显示以加速显示，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Draft**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例以草稿字体显示 Document1 窗口的内容。*/ Windows.Item("Document1").View.Draft = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例切换活动窗口的草稿字体选项。*/ ActiveDocument.ActiveWindow.View.Draft = !ActiveDocument.ActiveWindow.View.Draft` |

#### **View.FieldShading**

返回或设置域的屏幕底纹。**WdFieldShading** 类型，可读写。

**语法**

**express.FieldShading**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例激活活动窗口中所有窗体域的域底纹。*/ ActiveDocument.ActiveWindow.View.FieldShading = wdFieldShadingAlways` |

#### **View.FullScreen**

如果窗口为全屏显示，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.FullScreen**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口切换为全屏显示*/ Application.ActiveDocument.ActiveWindow.View.FullScreen = true  /*本示例激活 Sales.doc 窗口，并切换到非全屏显示*/ Application.Windows.Item("Sales.doc").Activate() Application.Windows.Item("Sales.doc").View.FullScreen = false` |

#### **View.Magnifier**

如果为 **True**，则在打印预览中将鼠标指针显示为放大镜形状，表示用户可以通过单击来放大页面的某个区域或通过缩小显示比例查看整页或展开多页。**Boolean** 类型，可读写。

**语法**

**express.Magnifier**

*express*   一个代表 **View** 对象的变量。

**说明**

如果不是在打印预览视图中，则该属性将产生错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将视图切换到打印预览，并把鼠标指针改为插入点。*/ function test() { PrintPreview = true ActiveDocument.ActiveWindow.View.Magnifier = false }` |

#### **View.MailMergeDataView**

如果在指定窗口中显示邮件合并数据，而不显示邮件合并域，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MailMergeDataView**

*express*   一个代表 **View** 对象的变量。

**说明**

如果指定窗口不是主控文档，则会产生一个错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档至少包含一个邮件合并域，则本示例显示附加数据源的第一个记录的邮件合并数据。*/ function test() { if(ActiveDocument.MailMerge.Fields.Count >= 1){     ActiveDocument.MailMerge.DataSource.ActiveRecord = 1     ActiveDocument.ActiveWindow.View.ShowFieldCodes = false     ActiveDocument.ActiveWindow.View.MailMergeDataView = true } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在查看邮件合并域和查看结果数据之间切换。*/ function test() { ActiveDocument.ActiveWindow.View.ShowFieldCodes = false ActiveDocument.ActiveWindow.View.MailMergeDataView = !ActiveDocument.ActiveWindow.View.MailMergeDataView }` |

#### **View.MarkupMode**

返回或设置一个 **WdRevisionsMode** 常量，该常量代表修订的显示模式。可读写。

**语法**

**express.MarkupMode**

*express*   一个代表 **View** 对象的变量。

#### **View.Panning**

返回或设置 **Boolean** 类型的值，该值表示 WPS 是否处于全景模式。可读写。

**语法**

**express.Panning**

*express*   一个代表 **View** 对象的变量。

**说明**

全景模式将鼠标指针变为手形，使用户能够通过按住鼠标来拖动文档，从而滚动查看文档页。

#### **View.Parent**

返回一个 **Object** 类型值，该值代表指定 **View** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **View** 对象的变量。

#### **View.ReadingLayout**

设置或返回一个 **Boolean** 类型的值，该值表示是否在阅读版式视图中查看文档。

**语法**

**express.ReadingLayout**

*express*   一个代表 **View** 对象的变量。

**说明**

如果为 **True**，则将文档切换到阅读版式视图。如果为 **False**，则关闭阅读版式视图。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭阅读版式视图。*/ ActiveDocument.ActiveWindow.View.ReadingLayout = false` |

#### **View.ReadingLayoutActualView**

设置或返回一个 **Boolean** 类型的值，该值代表是否使用与打印页面相同的版式在阅读版式视图中显示页面。

**语法**

**express.ReadingLayoutActualView**

*express*   一个代表 **View** 对象的变量。

**说明**

在阅读版式视图中，页面中不显示（能在普通视图和页面视图看到的）包含在打印页面中的全部内容。这些内容将在屏幕中予以显示。当 **ReadingLayoutActualView** 属性设为 **True** 时，文档显示将与打印出的页面一致。在较小的显示器上，会使用一定比例进行缩放，这将使文档不易阅读，在较大的显示器上显示效果就很好。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在阅读版式视图中显示的页面与打印出来的页面相同。*/ function test() { ActiveWindow.View.ReadingLayout = true ActiveWindow.View.ReadingLayoutActualView = true }  ` |

#### **View.ReadingLayoutAllowEditing**

返回 **Boolean** 值，该值表示是否允许在阅读布局模式下编辑文本。可读/写。

**语法**

**express.ReadingLayoutAllowEditing**

*express*   一个代表 **View** 对象的变量。

**说明**

此属性对应于阅读布局模式下**“视图选项”**菜单上的**“允许键入**/**禁止键入”**选项。

| 注释                                                   |
| ------------------------------------------------------ |
| 如果在阅读布局模式以外的其他视图中，则不能设置此属性。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例使 WPS 禁止在阅读布局模式下编辑文本。*/ ActiveDocument.ActiveWindow.View.ReadingLayoutAllowEditing = false` |

#### **View.ReadingLayoutAllowMultiplePages**

设置或返回一个 **Boolean** 类型的值，该值代表阅读版式视图是否并排显示两页。

**语法**

**express.ReadingLayoutAllowMultiplePages**

*express*   一个代表 **View** 对象的变量。

**说明**

WPS 可能允许、也可能不允许在阅读版式视图中显示两页。如果 WPS 不能保持合理的纵横比，则 WPS 将只显示一页。因此，如果将 **ReadingLayoutAllowMultiplePages** 属性设置为 **True** 并且 WPS 无法显示两页，则该属性将设为 **False**，并且 WPS 将仅显示单页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在版式允许的情况下在阅读版式视图中并排显示两页。*/ function test() { ActiveWindow.View.ReadingLayout = true ActiveWindow.View.ReadingLayoutAllowMultiplePages = true }` |

#### **View.ReadingLayoutTruncateMargins**

返回或设置一个 **WdReadingLayoutMargin** 常量，该常量代表当在阅读版式中查看文档时是显示还是隐藏边距。可读写。

**语法**

**express.ReadingLayoutTruncateMargins**

*express*   一个代表 **View** 对象的变量。

#### **View.Reviewers**

返回一个 **Reviewers** 对象，该对象代表所有审阅者。

**语法**

**express.Reviewers**

*express*   一个代表 **View** 对象的变量。

**说明**

**Reviewers** 对象是所有审阅者的全局列表，无论审阅者是否审阅指定窗口中显示的文档。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏文档中的所有修订和批注，只显示由“Jeff Smith”所作的修订和批注。*/ function HideRevisions(){     let view = ActiveWindow.View      view.ShowRevisionsAndComments = false     view.ShowFormatChanges = true     view.ShowInsertionsAndDeletions = true      for(let i = 1; i <= view.Reviewers.Count; i++){         view.Reviewers.Item(i).Visible = true     }      view.Reviewers.Item("Jeff Smith").Visible = true }` |

#### **View.RevisionsBalloonShowConnectingLines**

如果为 **True**，则 WPS 显示从文本到修订和批注气球之间的连接线。**Boolean** 类型，可读写。

**语法**

**express.RevisionsBalloonShowConnectingLines**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏文档文本和相应的修订或批注气球之间的连接线。*/ function ShowConnectingLines(){     ActiveWindow.View.RevisionsBalloonShowConnectingLines = false }` |

#### **View.RevisionsBalloonSide**

设置或返回一个 **WdRevisionsBalloonMargin** 常量，指定 WPS 是否在文档的左边距或右边距中显示修订气球。

**语法**

**express.RevisionsBalloonSide**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在左右边距之间切换修订气球。本示例假定活动窗口中的文档包含由一个或多个审阅者所做的修订，并且修订显示于气球中。*/ function ToggleRevisionBalloons(){     if(ActiveWindow.View.RevisionsBalloonSide == wdLeftMargin){         ActiveWindow.View.RevisionsBalloonSide = wdRightMargin     }     else{         ActiveWindow.View.RevisionsBalloonSide = wdLeftMargin     } }` |

#### **View.RevisionsBalloonWidth**

设置或返回一个 **Single** 类型的值，该值代表 WPS 中用于指定修订气球宽度的全局设置。可读写。

**语法**

**express.RevisionsBalloonWidth**

*express*   一个代表 **View** 对象的变量。

**说明**

修订气球的宽度包括文档边距与气球边缘之间 0.5 英寸宽的填充距离以及气球边缘与纸张边缘之间八分之一英寸的间距。 WPS 会在纸的左右边缘处添加空白。此宽度会扩展到边距，但并不更改文档的宽度或纸张大小。可以使用 **RevisionsBalloonWidthType** 属性指定在设置 **RevisionsBalloonWidth** 属性时的度量单位。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将修订气球的宽度设为 1 英寸并在左边距显示修订气球。本示例假定活动窗口中的文档包含由一个或多个审阅者所做的修订，并且修订显示于气球中。*/ function BalloonWidth(){     ActiveWindow.View.RevisionsBalloonWidthType = wdBalloonWidthPoints     ActiveWindow.View.RevisionsBalloonWidth = InchesToPoints(1)     ActiveWindow.View.RevisionsBalloonSide = wdLeftMargin }` |

#### **View.RevisionsBalloonWidthType**

设置或返回一个 **WdRevisionsBalloonWidthType** 常量，该常量代表指定 WPS 如何测量修订气球的宽度的全局设置。可读写。

**语法**

**express.RevisionsBalloonWidthType**

*express*   一个代表 **View** 对象的变量。

**说明**

**RevisionsBalloonWidthType** 属性设置用于设置 **RevisionsBalloonWidth** 属性的度量单位。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将修订气球的宽度设为文档宽度的 25%。本示例假定活动窗口中的文档包含多个审阅者所作的修订，并且修订显示于气球中。*/ function BalloonWidthType(){     ActiveWindow.View.RevisionsBalloonWidthType = wdBalloonWidthPercent     ActiveWindow.View.RevisionsBalloonWidth = 25 }` |

#### **View.RevisionsView**

设置或返回一个代表全局选项的 **WdRevisionsView** 常量，该选项用于指定 WPS 是显示文档的原始版本还是应用了修订和格式更改的版本。可读写。

**语法**

**express.RevisionsView**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在显示文档的原始版本和最终版本之间切换。本示例假定活动窗口中的文档包含由一个或多个审阅者所做的修订，并且修订显示于气球中。*/ function ToggleRevView(){     if(ActiveWindow.View.RevisionsMode == wdBalloonRevisions){         if(ActiveWindow.View.RevisionsView == wdRevisionsViewFinal){             ActiveWindow.View.RevisionsView = wdRevisionsViewOriginal         }         else{             ActiveWindow.View.RevisionsView = wdRevisionsViewFinal         }     } }` |

#### **View.SeekView**

返回或设置在页面视图中显示的文档元素。**WdSeekView** 类型，可读写。

**语法**

**express.SeekView**

*express*   一个代表 **View** 对象的变量。

**说明**

如果当前视图不是页面视图，该属性将产生一个错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档有脚注，则本示例在页面视图中显示它。*/ function test() { if(ActiveDocument.Footnotes.Count >= 1){     ActiveDocument.ActiveWindow.View.Type = wdPrintView     ActiveDocument.ActiveWindow.View.SeekView = wdSeekFootnotes } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示当前节中第一页的页脚。*/ function test() { ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = true ActiveDocument.ActiveWindow.View.Type = wdPrintView ActiveDocument.ActiveWindow.View.SeekView = wdSeekFirstPageFooter }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果在页面视图中选定脚注或尾注区域，本示例可切换到主控文档。*/ function test() { let myView = ActiveDocument.ActiveWindow.View  if(myView.SeekView == wdSeekFootnotes || myView.SeekView == wdSeekEndnotes){     myView.SeekView = wdSeekMainDocument } }` |

#### **View.ShadeEditableRanges**

返回或设置一个 **Long** 类型的值，该值代表是否将底纹应用于文档中用户有权修改的区域。

**语法**

**express.ShadeEditableRanges**

*express*   一个代表 **View** 对象的变量。

**说明**

如果为 **True**，则将底纹应用于文档中用户可以修改的区域。默认情况下，区域底纹处于打开状态。当区域底纹处于打开状态，或将该属性设置为 **True** 时，**ShadeEditableRanges** 属性返回的值为 65535。将 **ShadeEditableRanges** 属性设置为 **False** 时，返回的值为 0。该值仅用于代表属性值为 **True** 或 **False**，没有其他意义。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将底纹应用于用户有权修改的所有区域。*/ ActiveWindow.View.ShadeEditableRanges = true` |

#### **View.ShowAll**

如果显示所有非打印字符（如隐藏文字、制表符、空格和段落标记），则该属性值为 **True**。可读写 **Boolean** 类型。

**语法**

**express.ShowAll**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动窗口中显示所有非打印字符。*/ ActiveDocument.ActiveWindow.View.ShowAll = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在第一个窗口中切换非打印字符的显示方式。*/ Windows.Item(1).View.ShowAll = !Windows.Item(1).View.ShowAll` |

#### **View.ShowAnimation**

如果显示文本动画，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowAnimation**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例先激活活动窗口的文本动画，然后对选定内容应用闪烁文本动画。*/ function test() { ActiveDocument.ActiveWindow.View.ShowAnimation = true Selection.Font.Animation = wdAnimationSparkleText }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例关闭所有打开窗口中的字体动画。*/ function test() { for(let i = 1; i <= Windows.Count; i++){     Windows.Item(i).View.ShowAnimation = false } }` |

#### **View.ShowBookmarks**

如果在每个书签的首尾显示方括号，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowBookmarks**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在所有窗口中的书签周围显示方括号。*/ function test() { for(let i = 1; i <= Windows.Count; i++){     Windows.Item(i).View.ShowBookmarks = true } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用书签标记选定内容，在活动文档的每个书签周围显示方括号，然后折叠选定内容。*/ function test() { ActiveDocument.Bookmarks.Add("temp", Selection.Range) ActiveDocument.ActiveWindow.View.ShowBookmarks = true Selection.Collapse(wdCollapseStart) }` |

#### **View.ShowComments**

如果为 **True**，则 WPS 在文档中显示批注。**Boolean** 类型，可读写。

**语法**

**express.ShowComments**

*express*   一个代表 **View** 对象的变量。

**说明**

如果修订标记显示在右边距或左边距的气球中，则备注也显示在气球中。如果修订标记显示在文本中，则应用备注的文本被括在方括号中。当将鼠标指针置于括号中的文本时，相关的备注就会显示在鼠标指针上方的方形气球中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏活动文档中的所有备注。本示例假定活动窗口中的文档包括一个或多个备注。*/ function HideComments(){     ActiveWindow.View.ShowComments = false }` |

#### **View.ShowCropMarks**

返回或设置 **Boolean** 值，该值表示是否在页面的角显示裁剪标记，以指示页边距的位置。可读/写。

**语法**

**express.ShowCropMarks**

*express*   一个代表 **View** 对象的变量。

**说明**

如果显示裁减标记，将禁止用户通过拖动裁减标记来更改边距。显示裁减标记只是为了指示边距在页面上的位置。此属性对应于**“ WPS 选项”**对话框的**“高级”**选项卡中的**“裁减标记”**复选框。

| 注释                                                         |
| ------------------------------------------------------------ |
| 在东亚语言中默认显示裁减标记，在其他所有语言中默认情况下将关闭它。 |

#### **View.ShowDrawings**

如果在页面视图中显示由绘图工具创建的对象，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowDrawings**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口切换到页面视图，并显示由绘图工具创建的对象。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdPrintView ActiveDocument.ActiveWindow.View.ShowDrawings = true }  ` |

#### **View.ShowFieldCodes**

如果显示域代码，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowFieldCodes**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏 Document1 窗口的域代码。*/ Windows.Item("Document1").View.ShowFieldCodes = false` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示第一个窗口的域代码。*/ Windows.Item(1).View.ShowFieldCodes = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例切换活动窗口的域代码。*/ ActiveDocument.ActiveWindow.View.ShowFieldCodes = !ActiveDocument.ActiveWindow.View.ShowFieldCodes` |

#### **View.ShowFirstLineOnly**

如果在大纲视图中仅显示正文每段的第一行，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowFirstLineOnly**

*express*   一个代表 **View** 对象的变量。

**说明**

如果不是大纲或主控文档视图，该属性将产生一个错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口切换到大纲视图，并且隐藏除了正文每段的第一行之外所有的文本。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdOutlineView ActiveDocument.ActiveWindow.View.ShowFirstLineOnly = true }` |

#### **View.ShowFormat**

如果在大纲视图中显示字符格式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowFormat**

*express*   一个代表 **View** 对象的变量。

**说明**

如果不是大纲或主控文档视图，该属性将产生一个错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口切换到大纲视图，并显示字符格式。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdOutlineView ActiveDocument.ActiveWindow.View.ShowFormat = true }` |

#### **View.ShowFormatChanges**

如果为 **True**，则在启用“修订”功能后，WPS 显示对文档所作的格式更改。**Boolean** 类型，可读写。

**语法**

**express.ShowFormatChanges**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏对活动文档所作的格式更改。本示例假定在启用“修订”功能后，以对文档进行了格式更改。*/ function HideFormattingChanges(){     ActiveWindow.View.ShowFormatChanges = false }` |

#### **View.ShowHiddenText**

如果显示具有隐藏格式的文本，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowHiddenText**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏每个窗口中具有隐藏格式的文本。*/ function test() { for(let i = 1; i <= Windows.Count; i++){     Windows.Item(i).View.ShowHiddenText = false } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例切换对隐藏文本的显示。*/ ActiveDocument.ActiveWindow.View.ShowHiddenText = !ActiveDocument.ActiveWindow.View.ShowHiddenText` |

#### **View.ShowHighlight**

如果显示和打印文档中突出显示格式的文本，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowHighlight**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例切换对活动文档中突出显示文本的显示。*/ ActiveDocument.ActiveWindow.View.ShowHighlight = !ActiveDocument.ActiveWindow.View.ShowHighlight` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印活动文档，忽略突出显示格式。*/ function test() { ActiveDocument.ActiveWindow.View.ShowHighlight = false ActiveDocument.PrintOut() }` |

#### **View.ShowHyphens**

如果为 **True**，则显示可选连字符，该连字符指明在何处断开处于行尾的字。**Boolean** 类型，可读写。

**语法**

**express.ShowHyphens**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在选定内容之前插入一个可选连字符，然后在活动窗口显示可选连字符。*/ function test() { Selection.InsertBefore(String.fromCharCode(31)) ActiveDocument.ActiveWindow.View.ShowHyphens = true }` |

#### **View.ShowInkAnnotations**

返回或设置一个 **Boolean** 类型的值，该值表示显示还是隐藏手写墨迹注释。如果值为 **True**，则显示墨迹注释。如果值为 **False**，则隐藏墨迹注释。

**语法**

**express.ShowInkAnnotations**

*express*   一个代表 **View** 对象的变量。

**说明**

若要处理墨迹注释，必须在 Tablet 计算机上运行 WPS 。有关向文档添加手写墨迹注释的详细信息，请参阅“WPS 帮助”中的“使用墨迹注释标记文档”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档中的所有手写墨迹注释。*/ ActiveDocument.ActiveWindow.View.ShowInkAnnotations = true` |

#### **View.ShowInsertionsAndDeletions**

如果为 **True**，则 WPS 显示用“修订”功能在文档中插入和删除的文本。**Boolean** 类型，可读写。

**语法**

**express.ShowInsertionsAndDeletions**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏对文档所作的插入和删除修订。本示例假定活动窗口中的文档包含一个或多个审阅者所作的修订。*/ function HideInsertDelete(){     ActiveWindow.View.ShowInsertionsAndDeletions = false }` |

#### **View.ShowMainTextLayer**

如果在显示页眉和页脚区域时同时显示指定文档的文字，则该属性值为 **True**。该属性等效于**“页眉和页脚”**工具栏上的**“显示/隐藏文档文字”**按钮。**Boolean** 类型，可读写。

**语法**

**express.ShowMainTextLayer**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动窗口显示文档页眉并隐藏文档正文。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdPrintView ActiveDocument.ActiveWindow.View.SeekView = wdSeekCurrentPageHeader ActiveDocument.ActiveWindow.View.ShowMainTextLayer = false }` |

#### **View.ShowMarkupAreaHighlight**

返回或设置 **Boolean** 值，该值表示显示修订和注释气球的标记区域是否带底纹。可读/写。

**语法**

**express.ShowMarkupAreaHighlight**

*express*   一个代表 **View** 对象的变量。

**说明**

该属性对应于**“审阅”**功能区上**“显示标记”**菜单中的**“标记区域突出显示”**选项。

#### **View.ShowObjectAnchors**

如果在可在页面视图中定位的项目附近显示对象锁定标记，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowObjectAnchors**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例给选定内容添加一个框架，将活动窗口切换到页面视图，并显示框架对象的锁定标记。*/ function test() { Selection.Frames.Add(Selection.Range).LockAnchor = true ActiveDocument.ActiveWindow.View.Type = wdPrintView ActiveDocument.ActiveWindow.View.ShowObjectAnchors = true }  ` |

#### **View.ShowOptionalBreaks**

如果 WPS 显示可选换行符，则该属性值为 **True**。**Boolean** 类型，可读写

**语法**

**express.ShowOptionalBreaks**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动窗口中的可选换行符。*/ ActiveDocument.ActiveWindow.View.ShowOptionalBreaks = true` |

#### **View.ShowOtherAuthors**

如果应在文档中显示其他作者，则该属性值为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.ShowOtherAuthors**

*express*   一个代表 **View** 对象的变量。

**说明**

默认情况下，使用 WPS 2015 时，其他作者正在进行编辑或者已经从文档的某一部分隔离。在进行共同创作时使用 **ShowOtherAuthors** 属性可切换显示其他作者。

#### **View.ShowParagraphs**

如果显示段落标记，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowParagraphs**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏活动窗口中的段落标记。*/ ActiveDocument.ActiveWindow.View.ShowParagraphs = false` |

#### **View.ShowPicturePlaceHolders**

如果将空白方框作为图片的占位符进行显示，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowPicturePlaceHolders**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档中插入图片，并在活动窗口中显示图片占位符。*/ function test() { Selection.Collapse(wdCollapseStart) ActiveDocument.InlineShapes.AddPicture("C:\\anime.jpg", null, null, Selection.Range) ActiveDocument.ActiveWindow.View.ShowPicturePlaceHolders = true }  ` |

#### **View.ShowRevisionsAndComments**

如果为 **True**，则 WPS 显示使用“修订”功能对文档所作的修订和批注。**Boolean** 类型，可读写。

**语法**

**express.ShowRevisionsAndComments**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏文档中的修订和备注。本示例假定活动窗口中的文档包括一个或多个审阅者所作的修订。*/ function ShowRevsComments(){     ActiveWindow.View.ShowRevisionsAndComments = false }` |

#### **View.ShowSpaces**

如果显示空格字符，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowSpaces**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在选定内容之前插入空格，并显示活动窗口中的空格字符。*/ function test() { Selection.InsertBefore("    ") ActiveDocument.ActiveWindow.View.ShowSpaces = true }` |

#### **View.ShowTabs**

如果显示制表符，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowTabs**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在选定内容之前插入一个制表符，并显示 Document2 窗口中的制表符。*/ function test() { Windows.Item("Document2").Activate() Windows.Item("Document2").View.ShowTabs = true Selection.InsertBefore("\t") Selection.Collapse(wdCollapseEnd) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例拆分活动窗口，显示第一个窗格中的制表符，隐藏第二个窗格中的制表符。*/ function test() { ActiveDocument.ActiveWindow.Split = true ActiveDocument.ActiveWindow.Panes.Item(1).View.ShowTabs = true ActiveDocument.ActiveWindow.Panes.Item(2).View.ShowTabs = false }  ` |

#### **View.ShowTextBoundaries**

如果在页面视图中显示页边距、文本列、对象和框架周围的虚线，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowTextBoundaries**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口切换到页面视图，并显示文本边界线。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdPrintView ActiveDocument.ActiveWindow.View.ShowTextBoundaries = true }  ` |

#### **View.ShowXMLMarkup**

返回一个 **Long** 类型的值，该值表示是否在文档中显示 XML 标记。

**语法**

**express.ShowXMLMarkup**

*express*   一个代表 **View** 对象的变量。

**说明**

如果为 **True**，则表示显示标记。如果为 **False**，则表示隐藏标记。使用 **wdToggle**，可在显示和隐藏 XML 标记之间切换。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中切换显示和隐藏 XML 标记。*/ ActiveDocument.ActiveWindow.View.ShowXMLMarkup = wdToggle` |

#### **View.SplitSpecial**

返回或设置活动窗口的窗格。**WdSpecialPane** 类型，可读写。

**语法**

**express.SplitSpecial**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将在活动窗口的一个单独窗格中显示主页脚。*/ ActiveDocument.ActiveWindow.View.SplitSpecial = wdPanePrimaryFooter` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一个脚注添至活动文档，并在活动窗口的一个单独的窗格中显示所有脚注。*/ function test() { ActiveDocument.Footnotes.Add(Selection.Range,"1","Footnote text") ActiveDocument.ActiveWindow.View.Type = wdNormalView ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneFootnotes }  ` |

#### **View.TableGridlines**

如果显示表格虚框，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.TableGridlines**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动窗口中的表格虚框。*/ ActiveDocument.ActiveWindow.View.TableGridlines = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示与 Windows 集合中窗口 1 相关的各个窗格的表格虚框。*/ function test() { let myPanes = Windows.Item(1).Panes for(let i=1;i<=myPanes.Count;i++){     myPanes.Item(i).View.TableGridlines = true } }` |

#### **View.Type**

返回或设置视图的类型。可读/写 [WdViewType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.Type**

*express*   一个代表 **View** 对象的变量。

**说明**

**Type** 属性为当前处于大纲视图或主控文档视图的所有文档返回 [wdMasterView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)。除非预先在代码中明确设置，否则当前视图不会返回 **wdOutlineView**。

要检查当前文档是否为大纲视图，请使用 **Type** 属性和 **Subdocuments** 集合的 **Count** 属性。如果 **Type** 属性返回 [wdOutlineView ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)或 [wdMasterView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)，并且 **Count** 属性返回 0（零），则该文档为大纲视图。例如：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function VerifyOutlineView() {     if(Application.ActiveWindow.View.Type == wdOutlineView || Application.ActiveWindow.View.Type ==wdMasterView) {         if(Application.ActiveDocument.Subdocuments.Count == 0) {             MsgBox("当前是大纲视图")         }     } }` |

#### **View.WrapToWindow**

如果文字在文档窗口的右边缘换行，而不是在右边距或文本栏的右边界，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.WrapToWindow**

*express*   一个代表 **View** 对象的变量。

**说明**

该属性在页面视图或 Web 版式视图中无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使文本换行，以容纳在活动窗口中。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdNormalView ActiveDocument.ActiveWindow.View.WrapToWindow = true }` |

#### **View.Zoom**

返回一个 [Zoom ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Zoom/Zoom%20.htm#jsObject_Zoom)对象，该对象代表指定视图的显示比例。

**语法**

**express.Zoom**

*express*   一个代表 **View** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将所有打开窗口的显示比例更改为 125%*/ function wndBig(){     for(let i = 1;i <= Application.Windows.Count;i++){         Application.Windows.Item(i).View.Zoom.Percentage = 125     } }  /*本示例改变活动窗口的显示比例，以显示文本的全部宽度*/ Application.ActiveDocument.ActiveWindow.View.Zoom.PageFit = wdPageFitBestFit` |

适用环境：web

适用平台：windows/linux