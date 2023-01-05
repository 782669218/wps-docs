**SlideRange**



代表备注页或幻灯片范围的集合，该范围是一组幻灯片，少则仅包含一个，多则包含演示文稿中的所有幻灯片。

**说明**

您可以使用任何所需的幻灯片（从演示文稿的所有幻灯片中进行选择，或从选定内容的所有幻灯片中进行选择）构造幻灯片范围。例如，可以构造一个 **SlideRange** 集合，其中包含演示文稿中的前三张幻灯片、演示文稿中所有选定的幻灯片，或演示文稿中所有的标题幻灯片。

如同在用户界面中选中多个幻灯片并通过命令同时操作它们一样，通过建立一个 **SlideRange** 集合并对其使用属性和方法，可以在编程中同时操作多个幻灯片。如同用户界面中用于单张幻灯片的命令不能用于多张幻灯片一样，某些应用于单独 **SlideRange** 对象或只包含一张幻灯片的 **SlideRange** 集合的属性和方法不能用于包含多张幻灯片的 **SlideRange** 集合。一般情况下，如果选中多张幻灯片时，某些操作无法手动完成（例如返回某一幻灯片中的单个形状），则编程时也不能对包含多张幻灯片的 **SlideRange** 集合进行该操作。

对于用户界面中可用于一张或多张选中幻灯片的操作（例如复制幻灯片到剪贴板或设置幻灯片背景填充），相应的属性和方法也可用于包含多张幻灯片的 **SlideRange** 集合。下面是如何对多张幻灯片使用这些属性和方法的一些指导。

- 对 **SlideRange** 集合应用某方法等效于将该范围中所有 **Slide** 对象作为一个组来应用此方法。
- 设置 **SlideRange** 集合的属性值等效于逐一设置该范围中每个幻灯片的属性值（对于枚举类型的属性，设置“Mixed”值无效）。
- 如果 **SlideRange** 集合中的所有幻灯片对于某个返回枚举类型的属性具有相同的值，则该属性将返回该集合中单张幻灯片的属性值。如果该集合中的幻灯片对于该属性具有不同的值，则该属性将返回“Mixed”值。
- 如果 **SlideRange** 集合中的所有幻灯片对于某个返回简单数据类型（如 **Long**、**Single** 或 **String**）的属性具有相同的值，则该属性将返回该集合中单张幻灯片的该属性值。如果该集合中的幻灯片对于该属性具有不同的值，则该属性将返回 -2 或产生错误。例如，对于包含多张幻灯片的 **SlideRange** 对象使用 **Name** 属性将产生错误，因为每张幻灯片的 **Name** 属性都有不同的值。
- 幻灯片的某些格式属性不是通过直接应用于 **SlideRange** 集合的属性和方法来设置，而是通过应用于 **SlideRange** 集合中的对象（如 **ColorScheme** 对象）的属性和方法来设置。如果所含对象代表用户界面中可用于多个对象的操作，则可以从包含多张幻灯片的 **SlideRange** 集合返回该对象，且其属性和方法遵循前述规则。例如，可以使用 **ColorScheme** 属性返回 **ColorScheme** 对象，该对象代表用于指定 **SlideRange** 集合中所有幻灯片的配色方案。设置此 **ColorScheme** 对象的属性也将设置 **SlideRange** 集合中所有单张幻灯片的 **ColorScheme** 对象的相应属性。

以下示例说明如何执行下列操作：

- 返回一组通过名称或索引号指定的幻灯片
- 返回演示文稿中全部或部分选定的幻灯片
- 返回备注页
- 将属性和方法应用于某个幻灯片范围

备注：

对 **SlideRange** 对象代表的备注页使用下列属性和方法无效：**Copy** 方法、**Cut** 方法、**Delete** 方法、**Duplicate** 方法、**HeadersFooters** 属性、**Hyperlinks** 属性、**Layout** 属性、**PrintSteps** 属性、**SlideShowTransition** 属性。

**方法**

|                                                              | 名称                      | 说明                                                         |
| ------------------------------------------------------------ | ------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyTemplate**         | 将设计模板应用于指定的演示文稿。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyTheme**            | 对指定的幻灯片范围应用主题或设计模板。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyThemeColorScheme** | 对指定的幻灯片范围应用配色方案。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**                  | 将指定对象复制到剪贴板。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**                   | 删除指定对象并将其放到剪贴板中。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**                | 删除指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Duplicate**             | 创建指定的 **Slide** 或 **SlideRange** 对象的副本，将新的幻灯片或幻灯片范围添加到 **Slides** 集合中最初指定的幻灯片或幻灯片范围之后，然后返回代表复制幻灯片的 **Slide** 或 **SlideRange** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Export**                | 使用指定的图形筛选器导出幻灯片或幻灯片范围，并将导出的文件以指定的文件名保存。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**                  | 从指定集合中返回单个对象。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveTo**                | 将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**                | 选择指定的对象。                                             |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 返回一个 Application 对象，该对象表示指定对象的创建者。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Background**          | 返回代表幻灯片背景的 [ShapeRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange) 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorScheme**         | 返回或设置 ColorScheme 对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Comments**            | 返回一个 Comments 对象，该对象代表批注的集合。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**               | 返回指定集合中的对象数目。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Design**              | 返回代表设计的 Design 对象。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeadersFooters**      | 返回一个 HeadersFooters 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyperlinks**          | 返回一个代表指定幻灯片中的所有超链接的 Hyperlinks 集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Layout**              | 返回或设置一个 **PpSlideLayout** 常量，该常量代表幻灯片版式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Master**              | 返回一个 Master 对象，该对象代表幻灯片母版。只读。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                | 向一个演示文稿插入幻灯片时，Powerpoint 自动为幻灯片分配 Sliden 形式的名称，其中 n 是一个整数，代表幻灯片在演示文稿中创建的顺序。例如，在某个演示文稿中插入的第一张幻灯片被自动命名为 Slide1。如果将幻灯片从一个演示文稿复制到另一个演示文稿，则该幻灯片会失去它在第一个演示文稿中的名称并在第二个演示文稿中被自动分配新名称。幻灯片范围必须至少包含一张幻灯片。**String** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NotesPage**           | 返回一个 SlideRange 对象，该对象代表指定幻灯片或幻灯片区域的备注页。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shapes**              | 返回一个 Shapes 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideID**             | 返回指定幻灯片的唯一 ID 号。只读。**Long** 类型。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideIndex**          | 返回 **Slides** 集合内指定幻灯片的索引号。只读。**Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideNumber**         | 返回幻灯片编号。只读 Integer 类型。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideShowTransition** | 返回一个 SlideShowTransition 对象，该对象代表指定幻灯片切换的特殊效果。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThemeColorScheme**    | 返回一个 ThemeColorScheme 对象，该对象代表与指定幻灯片区域相关联的配色方案。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TimeLine**            | 返回 TimeLine 对象，该对象代表幻灯片的动画时间线。           |

**成员方法**

#### **SlideRange.ApplyTemplate**

将设计模板应用于指定的演示文稿。

**语法**

**express.ApplyTemplate(FileName)**

*express*   一个代表 **SlideRange** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**             |
| ---------- | ------------- | ------------ | -------------------- |
| *FileName* | 必选          | **String**   | 指定设计模板的名称。 |

**说明**

如果在字符串中引用未安装的演示文稿设计模板，则将产生运行时错误。无论 FeatureInstall 属性设置是什么，该模板都不自动安装。要对当前未安装的模板使用 ApplyTemplate 方法，必须先安装附加设计模板。为此，请通过运行 WPS Office安装程序（单击 Windows 控制面板中的“添加/删除程序”图标）安装WPP 的附加设计模板。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将“Professional”设计模板应用于当前演示文稿*/ Application.ActivePresentation.ApplyTemplate("c:\\program files\\office\\templates" + "\\presentation designs\\professional.pot") ` |

#### **SlideRange.ApplyTheme**

对指定的幻灯片范围应用主题或设计模板。

**语法**

**express.ApplyTheme(themeName)**

*express*   一个代表 **SlideRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *themeName* | 必选          | **String**   | 应用于 SlideRange 对象的主题文件 (.thmx) 或设计模板文件 (.pot) 的路径和名称。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对指定的幻灯片范围应用保存的主题*/ Application.ActivePresentation.Slides.Range([1, 3]).ApplyTheme("C:\\Program Files\\Office\\Templates\\MyTheme.thmx")  /*以下示例对指定的幻灯片范围应用保存的设计模板*/ Application.ActivePresentation.Slides.Range([1, 3]).ApplyTheme("C:\\Program Files\\Office\\Templates\\MyTheme.pot") ` |

#### **SlideRange.ApplyThemeColorScheme**

对指定的幻灯片范围应用配色方案。

**语法**

**express.ApplyThemeColorScheme(themeColorSchemeName)**

*express*   一个代表 **SlideRange** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型** | **说明**                                     |
| ---------------------- | ------------- | ------------ | -------------------------------------------- |
| *themeColorSchemeName* | 必选          | **String**   | 应用于幻灯片范围的配色方案文件的路径和名称。 |

#### **SlideRange.Copy**

将指定对象复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **SlideRange** 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在第一个窗口中拷贝选定幻灯片到剪贴板*/ Application.ActivePresentation.Windows.Item(1).Selection.Copy()  /*本示例将当前演示文稿中第一张幻灯片的第一个和第二个形状复制到剪贴板，然后将该副本粘贴到第二张灯片*/ function test(){ 	let slides2 = Application.ActivePresentation 	slides2.Slides.Item(1).Shapes.Range([1, 2]).Copy() 	slides2.Slides.Item(2).Shapes.Paste() }  /*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()` |

#### **SlideRange.Cut**

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut()**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中*/ Application.ActivePresentation.Windows.Item(1).Selection.Cut()  /*本示例从当前演示文稿中删除第一张幻灯片并将它的一个副本放到剪贴板中*/ Application.ActivePresentation.Slides.Item(1).Cut()  /*本示例删除当前演示文稿中第一张幻灯片的第一个形状中的文本，并将它的一个副本放到剪贴板*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Cut()` |

#### **SlideRange.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **SlideRange** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例删除选中的图形*/ Application.ActivePresentation.Windows.Item(1).Selection.Delete()` |

#### **SlideRange.Duplicate**

创建指定的 **Slide** 或 **SlideRange** 对象的副本，将新的幻灯片或幻灯片范围添加到 **Slides** 集合中最初指定的幻灯片或幻灯片范围之后，然后返回代表复制幻灯片的 **Slide** 或 **SlideRange** 对象。

**语法**

**express.Duplicate()**

*express*   一个代表 **SlideRange** 对象的变量。

#### **SlideRange.Export**

使用指定的图形筛选器导出幻灯片或幻灯片范围，并将导出的文件以指定的文件名保存。

**语法**

**express.Export(FileName, FilerName, ScaleWidth, ScaleHeight)**

*express*   一个代表 **SlideRange** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*    | 必选          | **String**   | 将导出并保存到磁盘的文件的名称。可以包括完整路径；如果不包括完整路径，WPP 就会在当前文件夹中创建一个文件。 |
| *FilerName*   | 必选          | **String**   | 要导出幻灯片的图形格式。指定的图形格式必须在 Windows 注册表中注册导出筛选器。可以指定注册的扩展名或注册的筛选器名称。WPP将首先在注册表中搜索匹配的扩展名。如果没有找到与指定字符串匹配的扩展名，WPP将查找匹配的筛选器名称。 |
| *ScaleWidth*  | 可选          | **Long**     | 导出幻灯片的宽度（以像素为单位）。                           |
| *ScaleHeight* | 可选          | **Long**     | 导出幻灯片的高度（以像素为单位）。                           |

**说明**

导出一个演示文稿不会将演示文稿的 Saved 属性设置为 True。
WPP使用指定的图形筛选器保存每张单独的幻灯片。导出并保存到磁盘的幻灯片的名称由WPP 决定。通常情况下，这些文件使用下列命名规范保存：Slide1.wmf、Slide2.wmf。保存文件的路径在 FileName 参数中指定。

#### **SlideRange.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **SlideRange** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Variant**  | 集合中要返回的单个对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例返回当前窗口的第二张幻灯片*/ Application.ActivePresentation.Slides.Item(2)` |

#### **SlideRange.MoveTo**

将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。

**语法**

**express.MoveTo(toPos)**

*express*   一个代表 **SlideRange** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *toPos*  | 必选          | **Long**     | 动画效果要移动到的索引。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在指定形状的动画效果集合中将一个动画效果移动到第二个动画效果的位置上*/ function test() {     let sldFirst = Application.ActivePresentation.Slides.Item(1)     let shpFirst = sldFirst.Shapes.Item(1)     let effAdd = sldFirst.TimeLine.MainSequence.AddEffect(shpFirst, msoAnimEffectBlinds)     effAdd.MoveTo(2) }  /*本示例将当前演示文稿中的第二张幻灯片移动到第一张幻灯片的位置上*/ Application.ActivePresentation.Slides.Item(2).MoveTo(1)` |

#### **SlideRange.Select**

选择指定的对象。

**语法**

**express.Select()**

*express*   一个代表 **SlideRange** 对象的变量。

**说明**

如果做出不适用于当前视图的选择，此代码将失败。例如，可以在幻灯片浏览视图中而不能在幻灯片视图中选择幻灯片。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择活动演示文稿中第一张幻灯片标题的前五个字符*/ Application.ActivePresentation.Slides.Item(1).Shapes.Title.TextFrame.TextRange.Characters(1, 5).Select()  /*本示例选择活动演示文稿中的第一张幻灯片*/ Application.ActivePresentation.Slides.Item(1).Select()  /*本示例选取一个已添加到演示文稿的新幻灯片中的表格。该表格具有三行和三列*/ Application.Presentations.Add().Slides.Add(1, ppLayoutBlank).Shapes.AddTable(3, 3).Select() ` |

**成员属性**

#### **SlideRange.Application**

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/ function test(pptPres) {     let pptPres = Application.ActivePresentation     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }   /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称*/ function test(){ 	for(let shpOle = 1; shpOle <= Application.ActivePresentation.Slides.Item(1).Shapes.Count; shpOle++) { 		if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(shpOle).Type == msoLinkedOLEObject) { 			MsgBox(Application.ActivePresentation.Slides.Item(1).Shapes.Item(shpOle).OLEFormat.Application.Name) 		} 	} }` |

#### **SlideRange.Background**

返回代表幻灯片背景的 [ShapeRange](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange) 对象。

**语法**

**express.Background**

*express*   一个代表 **SlideRange** 对象的变量。

**说明**

如果要使用 Background 属性设置单张幻灯片的背景而不更改幻灯片母版，则该幻灯片的 FollowMasterBackground 属性必须设置为 False。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将当前演示文稿的幻灯片母版的背景设为预设的底纹*/ Application.ActivePresentation.SlideMaster.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientLateSunset)   /*以下示例将当前演示文稿中第一张幻灯片的背景设为预设的底纹*/ function test(){ 	let slides2 =Application.ActivePresentation.Slides.Item(1) 	slides2.FollowMasterBackground = false 	slides2.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientLateSunset) }` |

#### **SlideRange.ColorScheme**

返回或设置 ColorScheme 对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。

**语法**

**express.ColorScheme**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将当前演示文稿第一张和第三张幻灯片的标题颜色设为绿色*/ function test(){     let mySlides = Application.ActivePresentation.Slides.Range([1, 3])     mySlides.ColorScheme.Colors(ppTitle).RGB = (0, 255, 0) }` |

#### **SlideRange.Comments**

返回一个 Comments 对象，该对象代表批注的集合。

**语法**

**express.Comments**

*express*   一个代表 **SlideRange** 对象的变量。

#### **SlideRange.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口*/ function test(){     let windows = Application.Windows     for(let i = 2; i <= windows2.Count; i++) {         Application.windows.Item(2).Close()     } }` |

#### **SlideRange.Design**

返回代表设计的 Design 对象。

**语法**

**express.Design**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回当前所在演示文档名*/ Application.ActivePresentation.Slides.Item(1).Design.Application.ActiveWindow ` |

#### **SlideRange.HeadersFooters**

返回一个 HeadersFooters 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。

**语法**

**express.HeadersFooters**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在当前演示文稿的备注母版中设置页脚文本*/ function test(){     let notesMaster = Application.ActivePresentation.NotesMaster.HeadersFooters         notesMaster.Footer.Text = "Regional Sales" }` |

#### **SlideRange.Hyperlinks**

返回一个代表指定幻灯片中的所有超链接的 Hyperlinks 集合。只读。

**语法**

**express.Hyperlinks**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例允许用户更新当前演示文稿中所有超链接中已过时的 Internet 地址*/ function test(){ 	let oldAddr =("Old Internet address") 	let newAddr = ("New Internet address")  	for(let s = 1; s <= Application.ActivePresentation.Slides.Count; s++) { 		for(let h = 1; h <= Application.ActivePresentation.Slides.Item(s).Hyperlinks.Count; h++) { 			if(Application.ActivePresentation.Slides.Item(s).Hyperlinks.Item(h).Address.toLowerCase() == oldAddr) { 				Application.ActivePresentation.Slides.Item(s).Hyperlinks.Item(h).Address = newAddr 			} 		} 	} }` |

#### **SlideRange.Layout**

返回或设置一个 **PpSlideLayout** 常量，该常量代表幻灯片版式。可读/写。

**语法**

**express.Layout**

*express*   一个代表 **SlideRange** 对象的变量。

**说明**

PpSlideLayout 可以是下列 PpSlideLayout 常量之一。
ppLayoutBlank
ppLayoutChart
ppLayoutChartAndText
ppLayoutClipartAndText
ppLayoutClipArtAndVerticalText
ppLayoutFourObjects
ppLayoutLargeObject
ppLayoutMediaClipAndText
ppLayoutMixed
ppLayoutObject
ppLayoutObjectAndText
ppLayoutObjectOverText
ppLayoutOrgchart
ppLayoutTable
ppLayoutText
ppLayoutTextAndChart
ppLayoutTextAndClipart
ppLayoutTextAndMediaClip
ppLayoutTextAndObject
ppLayoutTextAndTwoObjects
ppLayoutTextOverObject
ppLayoutTitle
ppLayoutTitleOnly
ppLayoutTwoColumnText
ppLayoutTwoObjectsAndText
ppLayoutTwoObjectsOverText
ppLayoutVerticalText
ppLayoutVerticalTitleAndText
ppLayoutVerticalTitleAndTextOverChart

#### **SlideRange.Master**

返回一个 Master 对象，该对象代表幻灯片母版。只读。

**语法**

**express.Master**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为当前演示文稿中第一张幻灯片的幻灯片母版设置背景填充*/ Application.ActivePresentation.Slides.Item(1).Master.Background.Fill.PresetGradient(msoGradientDiagonalUp, 1, msoGradientDaybreak) ` |

#### **SlideRange.Name**

向一个演示文稿插入幻灯片时，Powerpoint 自动为幻灯片分配 Sliden 形式的名称，其中 n 是一个整数，代表幻灯片在演示文稿中创建的顺序。例如，在某个演示文稿中插入的第一张幻灯片被自动命名为 Slide1。如果将幻灯片从一个演示文稿复制到另一个演示文稿，则该幻灯片会失去它在第一个演示文稿中的名称并在第二个演示文稿中被自动分配新名称。幻灯片范围必须至少包含一张幻灯片。**String** 类型，可读/写。

**语法**

**express.Name**

*express*   一个代表 **SlideRange** 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。例如，如果某个形状的 **Name** 属性值为“Rectangle 2”，则 .Shapes.Item("Rectangle 2") 将返回对该形状的引用。

#### **SlideRange.NotesPage**

返回一个 SlideRange 对象，该对象代表指定幻灯片或幻灯片区域的备注页。只读。

**语法**

**express.NotesPage**

*express*   一个代表 **SlideRange** 对象的变量。

**说明**

NotesPage 属性返回一张或一组幻灯片的备注页，且只允许对这些备注页进行修改。如果要进行会影响所有备注页的更改，请使用 NotesMaster 返回代表备注母版的 Slide 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为当前演示文稿中第一张幻灯片的备注页设置背景填充*/ function test(){     let myNotePage = Application.ActivePresentation.Slides.Item(1).NotesPage     myNotePage.FollowMasterBackground = false     myNotePage.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientLateSunset) } ` |

#### **SlideRange.Shapes**

返回一个 Shapes 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。

**语法**

**express.Shapes**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例添加一个宽 100 磅、高 50 磅的矩形，它的左上角距当前演示文稿第一张幻灯片的左边 5 磅、上边 25 磅*/ function test(){     let firstSlide = Application.ActivePresentation.Slides.Item(1)     firstSlide.Shapes.AddShape(msoShapeRectangle, 5, 25, 100, 50) }  /*以下示例设置当前演示文稿第一张幻灯片的第三个形状的填充纹理*/ function test(){     let newRect = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)     newRect.Fill.PresetTextured(msoTextureOak) }  /*假设当前演示文稿第一张幻灯片包含一个标题，以下示例的第二行和第三行设置该演示文稿第一张幻灯片的标题文本*/ function test(){     let newRect = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)     newRect.Fill.PresetTextured(msoTextureOak) }  /*假设当前演示文稿第二张幻灯片中的第二个形状包含文本框架，以下示例向该幻灯片添加一系列的段落。请注意：Chr(13) 用于在该文本中插入段落标记*/ function test(){     let tShape = Application.ActivePresentation.Slides.Item(2).Shapes.Item(2)     tShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item" }  /*对于大多数幻灯片版式，第一个形状为文本占位符。以下示例与上例完成相同功能*/ function test(){     let testShape = Application.ActivePresentation.Slides.Item(2).Shapes.Placeholders.Item(2)     testShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item" } ` |

#### **SlideRange.SlideID**

返回指定幻灯片的唯一 ID 号。只读。**Long** 类型。

**语法**

**express.SlideID**

*express*   一个代表 **SlideRange** 对象的变量。

**说明**

与 SlideIndex 属性不同，在演示文稿中添加或重新排列幻灯片时，Slide 对象的 SlideID 属性不会改变。因此，与使用具有幻灯片索引号的 Item 方法相比，使用具有幻灯片 ID 号的 FindBySlideID 方法是从 Slides 集合返回特定 Slide 对象的更可靠方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例示范添加并 Slide 检索 Slide 对象的唯一 ID 号*/ function test(){     let gslides = Application.ActivePresentation.Slides     // Get slide ID     let graphSlideID = gslides.Add(2, ppLayoutChart).SlideID }` |

#### **SlideRange.SlideIndex**

返回 **Slides** 集合内指定幻灯片的索引号。只读。**Long** 类型。

**语法**

**express.SlideIndex**

*express*   一个代表 **SlideRange** 对象的变量。

**说明**

与 SlideID 属性不同，在演示文稿中添加或重新排列幻灯片时，Slide 对象的 SlideIndex 属性会改变。因此，与使用具有幻灯片索引号的 Item 方法相比，使用具有幻灯片 ID 号的 FindBySlideID 方法是从 Slides 集合返回特定 Slide 对象的更可靠方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示第一个幻灯片放映窗口中当前放映幻灯片的索引号*/ Application.ActivePresentation.SlideShowWindow.View.Slide.SlideIndex` |

#### **SlideRange.SlideNumber**

返回幻灯片编号。只读 Integer 类型。

**语法**

**express.SlideNumber**

*express*   一个代表 **SlideRange** 对象的变量。

**说明**

Slide 对象的 SlideNumber 属性是显示幻灯片编号时在幻灯片右下角出现的实际编号。此编号由幻灯片在演示文稿中的编号（ SlideIndex 属性值）以及演示文稿的起始幻灯片编号（FirstSlideNumber 属性值）决定。幻灯片编号始终等于起始幻灯片编号 + 幻灯片索引号 - 1。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示第一张幻灯片编号的变化对某特定幻灯片编号的影响*/ function test(){ 	let mySlide = Application.ActivePresentation 	mySlide.PageSetup.FirstSlideNumber = 1     // starts slide numbering at 1 	MsgBox(mySlide.Slides.Item(2).SlideNumber) // returns 2  	mySlide.PageSetup.FirstSlideNumber = 10    // starts slide numbering at 10 	MsgBox(mySlide.Slides.Item(2).SlideNumber) // returns 11 }` |

#### **SlideRange.SlideShowTransition**

返回一个 SlideShowTransition 对象，该对象代表指定幻灯片切换的特殊效果。只读。

**语法**

**express.SlideShowTransition**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置在幻灯片放映过程中，当前演示文稿的第二张幻灯片在 5 秒后自动换片，并在幻灯片切换时播放犬吠声*/ function test(){ 	let transitionSet = Application.ActivePresentation.Slides.Item(2).SlideShowTransition 	transitionSet.AdvanceOnTime = true 	transitionSet.AdvanceTime = 5 	transitionSet.SoundEffect.ImportFromFile("C:\\windows\\media\\dogbark.wav")  	ActivePresentation.SlideShowSettings.AdvanceMode = ppSlideShowUseSlideTimings }` |

#### **SlideRange.ThemeColorScheme**

返回一个 ThemeColorScheme 对象，该对象代表与指定幻灯片区域相关联的配色方案。

**语法**

**express.ThemeColorScheme**

*express*   一个代表 **SlideRange** 对象的变量。

#### **SlideRange.TimeLine**

返回 TimeLine 对象，该对象代表幻灯片的动画时间线。

**语法**

**express.TimeLine**

*express*   一个代表 **SlideRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将一个跳动的动画添加到第一张幻灯片的第一个形状中*/ function test(){     let sldFirst = Application.ActivePresentation.Slides.Item(1)     let shpFirst = sldFirst.Shapes.Item(1)      sldFirst.TimeLine.MainSequence.AddEffect(shpFirst, msoAnimEffectBounce) }` |

适用环境：web

适用平台：windows/linux