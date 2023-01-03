**Slide**



代表一个幻灯片。**Slides** 集合包含演示文稿中的所有 **Slide** 对象。

**说明**

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif) 注释        |
| ------------------------------------------------------------ |
| 如果试图返回对单张幻灯片的引用却得到了一个 [SlideRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/SlideRange/SlideRange%20.htm#jsObject_SlideRange)对象，请不要奇怪。单张幻灯片既可以由 **Slide** 对象表示也可以由只包含一张幻灯片的 [SlideRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/SlideRange/SlideRange%20.htm#jsObject_SlideRange)集合表示，这取决于返回该幻灯片引用的方式。例如，如果使用 **A****dd** 方法创建并返回对幻灯片的引用，则该幻灯片由 **Slide** 对象表示。然而，如果使用 [Duplicate ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/SlideRange/SlideRange%20.htm#SlideRange.Duplicate)方法创建并返回对幻灯片的引用，则该幻灯片由包含单张幻灯片的 **SlideRange** 集合表示。因为应用于 **Slide** 对象的所有属性和方法也可应用于包含单张幻灯片的 **SlideRange** 集合，所以可对返回的幻灯片进行相同的操作，而不管它是由 **Slide** 对象还是 **SlideRange** 集合表示。 |

以下示例说明如何执行下列操作：

- 返回一个通过名称、索引号或幻灯片 ID 号指定的幻灯片
- 返回选定范围中的幻灯片
- 返回指定的任意文档窗口或幻灯片放映窗口中当前显示的幻灯片
- 新建幻灯片

## 示例

使用 **Slides**(*index*)（其中 *index* 是幻灯片名称或索引号）或者使用 **Slides.FindBySlideID**(*index*)（其中 *index* 是幻灯片 ID 号）返回单个 **Slide** 对象。以下示例设置活动演示文稿中第一张幻灯片的版式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Layout = ppLayoutTitle` |

使用 **Slides**(*index*)（其中 *index* 是幻灯片名称或索引号）或使用 **Slides.FindBySlideID**(*index*)（其中 *index* 是幻灯片 ID 号）返回单个 **Slide** 对象。以下示例设置活动演示文稿中第一张幻灯片的版式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Layout = ppLayoutTitle` |

以下示例设置 ID 号为 265 的幻灯片的版式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.FindBySlideID(265).Layout = ppLayoutTitle` |

使用 **Selection.SlideRange**(*index*) 返回单个 **Slide** 对象，其中 *index* 是选定范围中的幻灯片的名称或索引号。以下示例设置活动窗口选定范围中的第一张幻灯片的版式（假设至少选定一张幻灯片）。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveWindow.Selection.SlideRange.Item(1).Layout = ppLayoutTitle ` |

如果只选定了一张幻灯片，可以使用 **Selection.SlideRange** 返回包含选定幻灯片的 **SlideRange** 集合。以下示例设置当前窗口当前所选对象中第一张幻灯片的版式（假设正好只选定一张幻灯片）

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveWindow.Selection.SlideRange.Layout = ppLayoutTitle ` |

使用 **Slide** 属性返回指定文档窗口或幻灯片放映窗口视图中当前显示的幻灯片。以下示例将第二个文档窗口中当前显示的幻灯片复制到剪贴板。

| 示例代码复制                                    |
| ----------------------------------------------- |
| `Application.Windows.Item(2).View.Slide.Copy()` |

使用 **Add** 方法新建幻灯片并添加到演示文稿中。以下示例在当前演示文稿的开头添加一个标题幻灯片。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Add(1, ppLayoutTitleOnly)` |

**方法**

|                                                              | 名称                      | 说明                                                         |
| ------------------------------------------------------------ | ------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyTemplate**         | 将设计模板应用于指定的演示文稿。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyTheme**            | 对指定幻灯片应用主题或设计模板。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyThemeColorScheme** | 对指定幻灯片应用配色方案。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**                  | 将指定对象复制到剪贴板。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**                   | 删除指定对象并将其放到剪贴板中。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**                | 删除指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Export**                | 使用指定的图形筛选器导出幻灯片或幻灯片范围，并将导出的文件以指定的文件名保存。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveTo**                | 将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**                | 选择指定的对象。                                             |

**属性**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**          | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Background**           | 返回代表幻灯片背景的 **ShapeRange** 对象。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundStyle**      | 设置或返回指定对象的背景样式。可读/写。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorScheme**          | 返回或设置 **ColorScheme**对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Comments**             | 返回一个 **Comments** 对象，该对象代表批注的集合。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Design**               | 返回代表设计的 **Design** 对象。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeadersFooters**       | 返回一个 **HeadersFooters** 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HyperLinks**           | 返回一个代表指定幻灯片中的所有超链接的 **HyperLinks** 集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Layout**               | 返回或设置一个 **PpSlideLayout** 常量，该常量代表幻灯片版式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Master**               | 返回一个 **Master** 对象，该对象代表幻灯片母版。只读。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                 | 向一个演示文稿插入幻灯片时，Powerpoint 自动为幻灯片分配 Sliden 形式的名称，其中 n 是一个整数，代表幻灯片在演示文稿中创建的顺序。例如，在某个演示文稿中插入的第一张幻灯片被自动命名为 Slide1。如果将幻灯片从一个演示文稿复制到另一个演示文稿，则该幻灯片会失去它在第一个演示文稿中的名称并在第二个演示文稿中被自动分配新名称。幻灯片范围必须至少包含一张幻灯片。**String** 可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NotesPage**            | 返回一个 **SlideRange** 对象，该对象代表指定幻灯片或幻灯片区域的备注页。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shapes**               | 返回一个 [**Shapes**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#jsObject_Shapes) 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideID**              | 返回指定幻灯片的唯一 ID 号。只读。**Number** 类型。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideIndex**           | 返回 **Slides** 集合内指定幻灯片的索引号。只读。**Number** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideNumber**          | 返回幻灯片编号。只读 **Number** 类型。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlidesShowTransition** | 返回一个 **SlidesShowTransition** 对象，该对象代表指定幻灯片切换的特殊效果。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TimeLin**              | 返回 **TimeLine** 对象，该对象代表幻灯片的动画时间线。       |

**成员方法**

#### **Slide.ApplyTemplate**

将设计模板应用于指定的演示文稿。

**语法**

**express.ApplyTemplate(FileName)**

*express*   一个代表 **Slide** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**             |
| ---------- | ------------- | ------------ | -------------------- |
| *FileName* | 必选          | **String**   | 指定设计模板的名称。 |

**说明**

如果在字符串中引用未安装的演示文稿设计模板，则将产生运行时错误。无论 **Application.FeatureInstall** 属性设置是什么，该模板都不自动安装。要对当前未安装的模板使用 **ApplyTemplate** 方法，必须先安装附加设计模板。为此，请通过运行 WPS Office安装程序（单击 Windows 控制面板中的**“添加/删除程序”**图标）安装WPP 的附加设计模板。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将“Professional”设计模板应用于当前演示文稿。*/ Application.ActivePresentation.ApplyTemplate("c:\\program files\\office\\templates" + "\\presentationdesigns\\professional.pot")` |

#### **Slide.ApplyTheme**

对指定幻灯片应用主题或设计模板。

**语法**

**express.ApplyTheme(themeName)**

*express*   一个代表 **Slide** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *themeName* | 必选          | **String**   | 应用于 Slide 对象的主题文件 (.thmx) 或设计模板文件 (.pot) 的路径和名称。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对指定幻灯片应用保存的主题：*/ Application.ActivePresentation.Slides.Item(1).ApplyTheme("C:\\Program Files\\Office\\Templates\\MyTheme.thmx")  /*以下示例对指定幻灯片应用保存的设计模板：*/ Application.ActivePresentation.Slides.Item(1).ApplyTheme("C:\\Program Files\\Office\\Templates\\MyTheme.pot") ` |

#### **Slide.ApplyThemeColorScheme**

对指定幻灯片应用配色方案。

**语法**

**express.ApplyThemeColorScheme(themeColorSchemeName)**

*express*   一个代表 **Slide** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型** | **说明**                                 |
| ---------------------- | ------------- | ------------ | ---------------------------------------- |
| *themeColorSchemeName* | 必选          | **String**   | 应用于幻灯片的配色方案文件的路径和名称。 |

#### **Slide.Copy**

将指定对象复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **Slide** 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿的第一张幻灯片复制到剪贴板。*/ Application.ActivePresentation.Slides.Item(1).Copy()  /*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()  ` |

#### **Slide.Cut**

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut()**

*express*   一个代表 **Slide** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中。*/ Application.Windows.Item(1).Selection.Cut()  /*本示例从当前演示文稿中删除第一张幻灯片并将它的一个副本放到剪贴板中。*/ Application.ActivePresentation.Slides.Item(1).Cut()  /*本示例删除当前演示文稿中第一张幻灯片的第一个形状中的文本，并将它的一个副本放到剪贴板。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Cut() ` |

#### **Slide.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **Slide** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

#### **Slide.Export**

使用指定的图形筛选器导出幻灯片或幻灯片范围，并将导出的文件以指定的文件名保存。

**语法**

**express.Export(FileName, FilterName, ScaleWidth, ScaleHeight)**

*express*   一个代表 **Slide** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*    | 必选          | **String**   | 将导出并保存到磁盘的文件的名称。可以包括完整路径；如果不包括完整路径，WPP 就会在当前文件夹中创建一个文件。 |
| *FilterName*  | 必选          | **Single**   | 要导出幻灯片的图形格式。指定的图形格式必须在 Windows 注册表中注册导出筛选器。可以指定注册的扩展名或注册的筛选器名称。WPP将首先在注册表中搜索匹配的扩展名。如果没有找到与指定字符串匹配的扩展名，WPP将查找匹配的筛选器名称。 |
| *ScaleWidth*  | 可选          | **Number**   | 导出幻灯片的宽度（以像素为单位）。                           |
| *ScaleHeight* | 可选          | **Number**   | 导出幻灯片的高度（以像素为单位）。                           |

**说明**

导出一个演示文稿不会将演示文稿的 **Saved** 属性设置为 **True**。

WPP使用指定的图形筛选器保存每张单独的幻灯片。导出并保存到磁盘的幻灯片的名称由WPP 决定。通常情况下，这些文件保存为 Slide1.wmf、Slide2.wmf 等。保存文件的路径在 **FileName** 参数中指定。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动演示文稿中第三张幻灯片以 JPEG 图形格式导出到磁盘上。该幻灯片将保存为 Slide 3 of Annual Sales.jpg。*/ function test() {     let mySlide = Application.ActivePresentation.Slides.Item(3)     mySlide.Export("c:\\my documents\\Graphic Format\\" + "Slide 3 of Annual Sales", "JPG") }` |

#### **Slide.MoveTo**

将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。

**语法**

**express.MoveTo(toPos)**

*express*   一个代表 **Slide** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *toPos*  | 必选          | **Number**   | 动画效果要移动到的索引。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿中的第二张幻灯片移动到第一张幻灯片的位置上。*/ function test()  {     Application.ActivePresentation.Slides.Item(2).MoveTo(1) }` |

#### **Slide.Select**

选择指定的对象。

**语法**

**express.Select()**

*express*   一个代表 **Slide** 对象的变量。

**说明**

如果做出不适用于当前视图的选择，此代码将失败。例如，可以在幻灯片浏览视图中而不能在幻灯片视图中选择幻灯片

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择活动演示文稿中第一张幻灯片标题的前五个字符。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Title.TextFrame.TextRange.Characters(1, 5).Select()  /*本示例选择活动演示文稿中的第一张幻灯片。*/ Application.ActivePresentation.Slides.Item(1).Select()  /*本示例选取一个已添加到演示文稿的新幻灯片中的表格。该表格具有三行和三列。*/ function test() {     let mySlide = Application.Presentations.Add().Slides     mySlide.Add(1, ppLayoutBlank).Shapes.AddTable(3, 3).Select() }` |

**成员属性**

#### **Slide.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Slide** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(){     let pptPres = Application.ActivePresentation     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(s.Application.Path + "\\Added Slide") } /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {     for(let shpOle = 1; shpOle <= Application.ActivePresentation.Slides.Item(1).Shapes.Count; shpOle++)      {         if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(shpOle).Type == msoLinkedOLEObject)          {             alert(Application.ActivePresentation.Slides.Item(1).Shapes.Item(shpOle).OLEFormat.Application.Name)         }     } } ` |

#### **Slide.Background**

返回代表幻灯片背景的 **ShapeRange** 对象。

**语法**

**express.Background**

*express*   一个代表 **Slide** 对象的变量。

**说明**

如果要使用 **Background** 属性设置单张幻灯片的背景而不更改幻灯片母版，则该幻灯片的 **FollowMasterBackground** 属性必须设置为 **False**。

#### **Slide.BackgroundStyle**

设置或返回指定对象的背景样式。可读/写。

**语法**

**express.BackgroundStyle**

*express*   一个代表 **Slide** 对象的变量。

#### **Slide.ColorScheme**

返回或设置 **ColorScheme**对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。

**语法**

**express.ColorScheme**

*express*   一个代表 **Slide** 对象的变量。

#### **Slide.Comments**

返回一个 **Comments** 对象，该对象代表批注的集合。

**语法**

**express.Comments**

*express*   一个代表 **Slide** 对象的变量。

#### **Slide.Design**

返回代表设计的 **Design** 对象。

**语法**

**express.Design**

*express*   一个代表 **Slide** 对象的变量。

#### **Slide.HeadersFooters**

返回一个 **HeadersFooters** 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。

**语法**

**express.HeadersFooters**

*express*   一个代表 **Slide** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在当前演示文稿的备注母版中设置页脚文本、日期和时间的格式，并将日期和时间设定为可自动修改。*/ function test() {     let notesmaster = Application.ActivePresentation.NotesMaster.HeadersFooters     notesmaster.Footer.Text = "Regional Sales"     let newnotesmaster = notesmaster2.DateAndTime     newnotesmaster.UseFormat = true     newnotesmaster.Format = ppDateTimeHmmss }` |

#### **Slide.HyperLinks**

返回一个代表指定幻灯片中的所有超链接的 **HyperLinks** 集合。只读。

**语法**

**express.HyperLinks**

*express*   一个代表 **Slide** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例允许用户更新当前演示文稿中所有超链接中已过时的 Internet 地址。*/ function test() {     let oldAddr = "Old Internet address"     let newAddr = "New Internet address"     for(let s = 1; s <= Application.ActivePresentation.Slides.Count; s++)      {         for(let h = 1; h <= Application.ActivePresentation.Slides.Item(s).Hyperlinks.Count; h++)          {             if(Application.ActivePresentation.Slides.Item(s).Hyperlinks.Item(h).Address.toLowerCase() == oldAddr)              {                 Application.ActivePresentation.Slides.Item(s).Hyperlinks.Item(h).Address = newAddr             }         }     } }` |

#### **Slide.Layout**

返回或设置一个 **PpSlideLayout** 常量，该常量代表幻灯片版式。可读/写。

**语法**

**express.Layout**

*express*   一个代表 **Slide** 对象的变量。

**说明**

| PpSlideLayout 可以是下列 PpSlideLayout 常量之一。 |
| ------------------------------------------------- |
| **ppLayoutBlank**                                 |
| **ppLayoutChart**                                 |
| **ppLayoutChartAndText**                          |
| **ppLayoutClipartAndText**                        |
| **ppLayoutClipArtAndVerticalText**                |
| **ppLayoutFourObjects**                           |
| **ppLayoutLargeObject**                           |
| **ppLayoutMediaClipAndText**                      |
| **ppLayoutMixed**                                 |
| **ppLayoutObject**                                |
| **ppLayoutObjectAndText**                         |
| **ppLayoutObjectOverText**                        |
| **ppLayoutOrgchart**                              |
| **ppLayoutTable**                                 |
| **ppLayoutText**                                  |
| **ppLayoutTextAndChart**                          |
| **ppLayoutTextAndClipart**                        |
| **ppLayoutTextAndMediaClip**                      |
| **ppLayoutTextAndObject**                         |
| **ppLayoutTextAndTwoObjects**                     |
| **ppLayoutTextOverObject**                        |
| **ppLayoutTitle**                                 |
| **ppLayoutTitleOnly**                             |
| **ppLayoutTwoColumnText**                         |
| **ppLayoutTwoObjectsAndText**                     |
| **ppLayoutTwoObjectsOverText**                    |
| **ppLayoutVerticalText**                          |
| **ppLayoutVerticalTitleAndText**                  |
| **ppLayoutVerticalTitleAndTextOverChart**         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例中，如果当前演示文稿的第一张幻灯片原来只有一个标题，则改变该幻灯片的版式以包含一个标题和一个副标题。*/ function test() {     let slides = Application.ActivePresentation.Slides.Item(1)     if(slides.Layout == ppLayoutTitleOnly)      {         slides.Layout = ppLayoutTitle     } }` |

#### **Slide.Master**

返回一个 **Master** 对象，该对象代表幻灯片母版。只读。

**语法**

**express.Master**

*express*   一个代表 **Slide** 对象的变量。

#### **Slide.Name**

向一个演示文稿插入幻灯片时，Powerpoint 自动为幻灯片分配 Sliden 形式的名称，其中 n 是一个整数，代表幻灯片在演示文稿中创建的顺序。例如，在某个演示文稿中插入的第一张幻灯片被自动命名为 Slide1。如果将幻灯片从一个演示文稿复制到另一个演示文稿，则该幻灯片会失去它在第一个演示文稿中的名称并在第二个演示文稿中被自动分配新名称。幻灯片范围必须至少包含一张幻灯片。**String** 可读/写。

**语法**

**express.Name**

*express*   一个代表 **Slide** 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。

#### **Slide.NotesPage**

返回一个 **SlideRange** 对象，该对象代表指定幻灯片或幻灯片区域的备注页。只读。

**语法**

**express.NotesPage**

*express*   一个代表 **Slide** 对象的变量。

**说明**

对 **SlideRange** 对象代表的备注页使用下列属性和方法无效：**Copy** 方法、**Cut** 方法、**Delete** 方法、**Duplicate** 方法、**HeadersFooters** 属性、**Hyperlinks** 属性、**Layout** 属性、**PrintSteps** 属性、**SlideShowTransition** 属性。**NotesPage** 属性返回一张或一组幻灯片的备注页，且只允许对这些备注页进行修改。如果要进行会影响所有备注页的更改，请使用 **NotesMaster** 返回代表备注母版的 **Slide** 对象。

#### **Slide.Shapes**

返回一个 [**Shapes**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#jsObject_Shapes) 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。

**语法**

**express.Shapes**

*express*   一个代表 **Slide** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例添加一个宽 100 磅、高 50 磅的矩形，它的左上角距当前演示文稿第一张幻灯片的左边 5 磅、上边 25 磅。*/ function test() {     let firstSlide = Application.ActivePresentation.Slides.Item(1)     firstSlide.Shapes.AddShape(msoShapeRectangle, 5, 25, 100, 50) }  /*以下示例设置当前演示文稿第一张幻灯片的第三个形状的填充纹理。*/ function test() {     let newRect = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)     newRect.Fill.PresetTextured(msoTextureOak) }  /*假设当前演示文稿第一张幻灯片包含一个标题，以下示例的第二行和第三行设置该演示文稿第一张幻灯片的标题文本。*/ function test() {     let firstSl = Application.ActivePresentation.Slides.Item(1)     firstSl.Shapes.Title.TextFrame.TextRange.Text = "Some title text"     firstSl.Shapes.Item(1).TextFrame.TextRange.Text = "Other title text" }  /*假设当前演示文稿第二张幻灯片中的第二个形状包含文本框架，以下示例向该幻灯片添加一系列的段落。请注意：Chr(13) 用于在该文本中插入段落标记。* function test() {     let tShape = Application.ActivePresentation.Slides.Item(2).Shapes.Item(2)     tShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item" }  /*对于大多数幻灯片版式，第一个形状为文本占位符。以下示例与上例完成相同功能。*/ function test() {     let testShape = Application.ActivePresentation.Slides.Item(2).Shapes.Placeholders.Item(2)     testShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item"     }` |

#### **Slide.SlideID**

返回指定幻灯片的唯一 ID 号。只读。**Number** 类型。

**语法**

**express.SlideID**

*express*   一个代表 **Slide** 对象的变量。

**说明**

与 SlideIndex 属性不同，在演示文稿中添加或重新排列幻灯片时，Slide 对象的 SlideID 属性不会改变。因此，与使用具有幻灯片索引号的 Item 方法相比，使用具有幻灯片 ID 号的 FindBySlideID 方法是从 Slides 集合返回特定 Slide 对象的更可靠方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例示范如何检索 Slide 对象的唯一 ID 号，并使用此编号从 Slide 集合返回该 Slides 对象。*/ function test() {     let gslides = Application.ActivePresentation.Slides      //Get slide ID     let graphSlideID = gslides.Add(2, ppLayoutChart).SlideID     gslides.FindBySlideID(graphSlideID).SlideShowTransition.EntryEffect = ppEffectCoverLeft     //Use ID to return specific slide }` |

#### **Slide.SlideIndex**

返回 **Slides** 集合内指定幻灯片的索引号。只读。**Number** 类型。

**语法**

**express.SlideIndex**

*express*   一个代表 **Slide** 对象的变量。

**说明**

与 SlideID 属性不同，在演示文稿中添加或重新排列幻灯片时，Slide 对象的 SlideIndex 属性会改变。因此，与使用具有幻灯片索引号的 Item 方法相比，使用具有幻灯片 ID 号的 **FindBySlideID** 方法是从 Slides 集合返回特定 Slide 对象的更可靠方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示第一个幻灯片放映窗口中当前放映幻灯片的索引号。*/ alert(Application.SlideShowWindows.Item(1).View.Slide.SlideIndex)` |

#### **Slide.SlideNumber**

返回幻灯片编号。只读 **Number** 类型。

**语法**

**express.SlideNumber**

*express*   一个代表 **Slide** 对象的变量。

**说明**

Slide 对象的 SlideNumber 属性是显示幻灯片编号时在幻灯片右下角出现的实际编号。此编号由幻灯片在演示文稿中的编号（ SlideIndex 属性值）以及演示文稿的起始幻灯片编号（FirstSlideNumber 属性值）决定。幻灯片编号始终等于起始幻灯片编号 + 幻灯片索引号 - 1。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示第一张幻灯片编号的变化对某特定幻灯片编号的影响。*/ function test() {     let slides2 = Application.ActivePresentation     slides2.PageSetup.FirstSlideNumber = 1        //starts slide numbering at 1     alert(slides2.Slides.Item(2).SlideNumber)    //returns 2      slides2.PageSetup.FirstSlideNumber = 10       //starts slide numbering at 10     alert(slides2.Slides.Item(2).SlideNumber)    //returns 11 }` |

#### **Slide.SlidesShowTransition**

返回一个 **SlidesShowTransition** 对象，该对象代表指定幻灯片切换的特殊效果。只读。

**语法**

**express.SlidesShowTransition**

*express*   一个代表 **Slide** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置在幻灯片放映过程中，当前演示文稿的第二张幻灯片在 5 秒后自动换片，并在幻灯片切换时播放犬吠声。*/ function test() {     let slides2 = Application.ActivePresentation.Slides.Item(2).SlideShowTransition     slides2.AdvanceOnTime = true     slides2.AdvanceTime = 5     slides2.SoundEffect.ImportFromFile("c:\\windows\\media\\dogbark.wav")      Application.ActivePresentation.SlideShowSettings.AdvanceMode = ppSlideShowUseSlideTimings }` |

#### **Slide.TimeLin**

返回 **TimeLine** 对象，该对象代表幻灯片的动画时间线。

**语法**

**express.TimeLin**

*express*   一个代表 **Slide** 对象的变量。

适用环境：web

适用平台：windows/linux