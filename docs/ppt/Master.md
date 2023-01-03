**Master**



代表一个幻灯片母版、标题母版、讲义母版、备注母版或设计母版。

**说明**

若要返回一个 **Master** 对象，请使用 **Slide**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Slide/Slide%20.htm#jsObject_Slide)对象或 **SlideRange** 集合的 **Master**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/SlideRange/SlideRange%20.htm#SlideRange.Master)属性，或使用 **Presentation**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#jsObject_Presentation)对象的 **HandoutMaster**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#Presentation.HandoutMaster)、**NotesMaster**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#Presentation.NotesMaster)、**SlideMaster**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#Presentation.SlideMaster)或 **TitleMaster**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#Presentation.TitleMaster)属性。请注意，这些属性中的某些也可用于 **Design** 对象。以下示例设置活动演示文稿的幻灯片母版的背景填充。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.SlideMaster.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientBrass)` |

若要为演示文稿添加标题母版或设计并返回代表新标题母版或设计的 

[Master](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Master/Master%20.htm#jsObject_Master)

[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Master/Master%20.htm#jsObject_Master)对象，请使用 

AddTitleMaster

 属性。以下示例在活动演示文稿中添加一个标题母版，并将标题占位符放置在距母版顶部 10 磅的位置。

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.AddTitleMaster.Shapes.Title.Top = 10` |

**方法**

|                                                              | 名称           | 说明                                                         |
| ------------------------------------------------------------ | -------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyTheme** | 将主题或设计模板应用于指定的幻灯片母板、标题母版、讲义母版、备注母版或设计母版。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**     | 删除指定的对象。                                             |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Background**          | 返回代表幻灯片背景的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundStyle**     | 设置或返回指定对象的背景样式。可读/写。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorScheme**         | 返回或设置 **ColorScheme** 对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Design**              | 返回代表设计的 **Design** 对象。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeadersFooters**      | 返回一个 **HeadersFooters** 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**              | 以磅为单位返回或设置指定对象的高度。用于 **Master** 对象时只读，**Single** 类型；用于所有其他对象时可读/写，**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyperlinks**          | 返回一个代表指定幻灯片中的所有超链接的 **Hyperlinks** 集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                | 返回或设置指定对象的名称。**String** 类型，可读/写。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shapes**              | 返回一个 **Shapes** 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideShowTransition** | 返回一个 **SlideShowTransition** 对象，该对象代表指定幻灯片切换的特殊效果。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextStyles**          | 返回一个 **TextStyles** 集合，该集合代表用于指定幻灯片母版的三种文本样式 - 标题文本、正文文本和默认文本。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Theme**               | 返回一个 **Theme** 对象，该对象代表指定的幻灯片母版、标题母版、讲义母版、备注母版或设计母版所使用的主题。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TimeLine**            | 返回 **TimeLine** 对象，该对象代表幻灯片的动画时间线。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**               | 以磅为单位返回或设置指定对象的宽度。用于 **Master**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Master/Master%20.htm#jsObject_Master)对象时只读，**Single** 类型；用于所有其他对象时可读/写，**Single** 类型。 |

**成员方法**

#### **Master.ApplyTheme**

将主题或设计模板应用于指定的幻灯片母板、标题母版、讲义母版、备注母版或设计母版。

**语法**

**express.ApplyTheme(themeName )**

*express*   一个代表 **Master** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *themeName* | 必选          | **String**   | 应用于 **Master** 对象的主题文件 (.thmx) 或设计模板文件 (.pot) 的路径和名称。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将保存的主题应用于幻灯片母板：*/ Application.ActivePresentation.SlideMaster.ApplyTheme("C:\\Program Files\\ Office\\Templates\\MyTheme.thmx")  /*以下示例将保存的设计模板应用于幻灯片母板：*/ Application.ActivePresentation.SlideMaster.ApplyTheme("C:\\Program Files\\ Office\\Templates\\MyTheme.pot") ` |

#### **Master.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **Master** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**成员属性**

#### **Master.Application**

返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Master** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(pptPres){     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **Master.Background**

返回代表幻灯片背景的 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。 

**语法**

**express.Background**

*express*   一个代表 **Master** 对象的变量。

**说明**

如果要使用 **Background** 属性设置单张幻灯片的背景而不更改幻灯片母版，则该幻灯片的 **FollowMasterBackground** 属性必须设置为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将当前演示文稿的幻灯片母版的背景设为预设的底纹。*/ Application.ActivePresentation.SlideMaster.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientLateSunset)  /*以下示例将当前演示文稿中第一张幻灯片的背景设为预设的底纹。*/ function test() {     let sli = Application.ActivePresentation.Slides.Item(1)     sli.FollowMasterBackground = false     sli.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientLateSunset) }` |

#### **Master.BackgroundStyle**

设置或返回指定对象的背景样式。可读/写。

**语法**

**express.BackgroundStyle**

*express*   一个代表 **Master** 对象的变量。

#### **Master.ColorScheme**

返回或设置 **ColorScheme** 对象，该对象代表指定幻灯片、幻灯片区域或幻灯片母版的配色方案。可读/写。

**语法**

**express.ColorScheme**

*express*   一个代表 **Master** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将当前演示文稿第一张和第三张幻灯片的标题颜色设为绿色。*/ function test() {   let mySlides = Application.ActivePresentation.Slides.Range(Array(1, 3))   mySlides.ColorScheme.Colors(ppTitle).RGB = (0, 255, 0) }` |

#### **Master.Design**

返回代表设计的 **Design** 对象。

**语法**

**express.Design**

*express*   一个代表 **Master** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例添加一个标题母版。*/ function test(){     Application.ActivePresentation.Slides.Item(1).Design.AddTitleMaster() }` |

#### **Master.HeadersFooters**

返回一个 **HeadersFooters** 集合，该集合代表与幻灯片、幻灯片母版或幻灯片区域相关联的页眉、页脚、日期和时间以及幻灯片编号。只读。

**语法**

**express.HeadersFooters**

*express*   一个代表 **Master** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `以下示例在当前演示文稿的备注母版中设置页脚文本、日期和时间的格式，并将日期和时间设定为可自动修改。 function test() {   let hea = Application.ActivePresentation.NotesMaster.HeadersFooters   hea.Footer.Text = "Regional Sales"   let tim = hea.DateAndTime   tim.UseFormat = true   tim.Format = ppDateTimeHmmss }` |

#### **Master.Height**

以磅为单位返回或设置指定对象的高度。用于 **Master** 对象时只读，**Single** 类型；用于所有其他对象时可读/写，**Single** 类型。

**语法**

**express.Height**

*express*   一个代表 **Master** 对象的变量。

**说明**

一个 **Height****Shape** 对象的 **Shape****Height** 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将第二个文档窗口的高度设置为应用程序窗口高度的一半。*/ Application.Windows.Item(2).Height = Application.Height / 2` |

#### **Master.Hyperlinks**

返回一个代表指定幻灯片中的所有超链接的 **Hyperlinks** 集合。只读。

**语法**

**express.Hyperlinks**

*express*   一个代表 **Master** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例允许用户更新当前演示文稿中所有超链接中已过时的 Internet 地址。*/ function test(){   let oldAddr = Application.prompt("Old Internet address")   let newAddr = Application.prompt("New Internet address")       for(let i = 1; i <= Application.ActivePresentation.Slides.Count; i++){           for(let ii = 1; ii <= Application.ActivePresentation.Slides.Item(i).Hyperlinks.Count; ii++){               if(LCase(Application.ActivePresentation.Slides.Item(i).Hyperlinks.Item(ii).Address) == oldAddr){                   Application.ActivePresentation.Slides.Item(i).Hyperlinks.Item(ii).Address = newAddr           }        }   } }` |

#### **Master.Name**

返回或设置指定对象的名称。**String** 类型，可读/写。

**语法**

**express.Name**

*express*   一个代表 **Master** 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。例如，如果某个形状的 **Name** 属性值为“Rectangle 2”，则 .Shapes("Rectangle 2") 将返回对该形状的引用。

#### **Master.Shapes**

返回一个 **Shapes** 集合，该集合代表被放置或插入到指定幻灯片、幻灯片母版或幻灯片组的所有元素。该集合可以包含绘图、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号以及日期和时间对象，这些对象位于幻灯片上或备注页中的幻灯片映像上。只读。

**语法**

**express.Shapes**

*express*   一个代表 **Master** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例添加一个宽 100 磅、高 50 磅的矩形，它的左上角距当前演示文稿第一张幻灯片的左边 5 磅、上边 25 磅。*/ function test() {   let firstSlide = Application.ActivePresentation.Slides.Item(1)   firstSlide.Shapes.AddShape(msoShapeRectangle, 5, 25, 100, 50) }  /*以下示例设置当前演示文稿第一张幻灯片的第三个形状的填充纹理。*/ function test() {   let newRect = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)   newRect.Fill.PresetTextured(msoTextureOak) }  /*假设当前演示文稿第一张幻灯片包含一个标题，以下示例的第二行和第三行设置该演示文稿第一张幻灯片的标题文本。*/ function test() {   let firstSl = Application.ActivePresentation.Slides.Item(1)   firstSl.Shapes.Title.TextFrame.TextRange.Text = "Some title text"   firstSl.Shapes.Item(1).TextFrame.TextRange.Text = "Other title text" }  /*假设当前演示文稿第二张幻灯片中的第二个形状包含文本框架，以下示例向该幻灯片添加一系列的段落。请注意：Chr(13) 用于在该文本中插入段落标记。*/ function test() {   let tShape = Application.ActivePresentation.Slides.Item(2).Shapes.Item(2)   tShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item" }  /*对于大多数幻灯片版式，第一个形状为文本占位符。以下示例与上例完成相同功能。*/ function test() {   let testShape = Application.ActivePresentation.Slides.Item(2).Shapes.Placeholders.Item(2)   testShape.TextFrame.TextRange.Text = "First Item" + "\r" + "Second Item" + "\r" + "Third Item" }` |

#### **Master.SlideShowTransition**

返回一个 **SlideShowTransition** 对象，该对象代表指定幻灯片切换的特殊效果。只读。

**语法**

**express.SlideShowTransition**

*express*   一个代表 **Master** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `以下示例设置在幻灯片放映过程中，当前演示文稿的第二张幻灯片在 5 秒后自动换片，并在幻灯片切换时播放犬吠声 function test() {   let sli = Application.ActivePresentation.Slides.Item(2).SlideShowTransition   sli.AdvanceOnTime = true   sli.AdvanceTime = 5   sli.SoundEffect.ImportFromFile("c:\\windows\\media\\dogbark.wav")   Application.ActivePresentation.SlideShowSettings.AdvanceMode = ppSlideShowUseSlideTimings }` |

#### **Master.TextStyles**

返回一个 **TextStyles** 集合，该集合代表用于指定幻灯片母版的三种文本样式 - 标题文本、正文文本和默认文本。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**语法**

**express.TextStyles**

*express*   一个代表 **Master** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例为活动演示文稿中幻灯片上的第一级正文文本设置字体名称和字号。 function test() {   let sli = Application.ActivePresentation.SlideMaster.TextStyles.Item(ppBodyStyle).Levels.Item(1)   let fon = sli.Font   fon.Name = "arial"   fon.Size = 36 }` |

#### **Master.Theme**

返回一个 **Theme** 对象，该对象代表指定的幻灯片母版、标题母版、讲义母版、备注母版或设计母版所使用的主题。

**语法**

**express.Theme**

*express*   一个代表 **Master** 对象的变量。

#### **Master.TimeLine**

返回 **TimeLine** 对象，该对象代表幻灯片的动画时间线。

**语法**

**express.TimeLine**

*express*   一个代表 **Master** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将一个跳动的动画添加到第一张幻灯片的第一个形状中。*/ function test(){     let sldFirst = Application.ActivePresentation.Slides.Item(1)     let shpFirst = sldFirst.Shapes.Item(1)     sldFirst.TimeLine.MainSequence.AddEffect(shpFirst, msoAnimEffectBounce) }` |

#### **Master.Width**

以磅为单位返回或设置指定对象的宽度。用于 **Master**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Master/Master%20.htm#jsObject_Master)对象时只读，**Single** 类型；用于所有其他对象时可读/写，**Single** 类型。

**语法**

**express.Width**

*express*   一个代表 **Master** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所有水平可用空间。要使以下示例执行，只能有两个打开的文档窗口。*/ function test() {   let ah = Application.Windows.Item(1).Height                           // available height   let aw = Application.Windows.Item(1).Width + Windows.Item(2).Width    // available width   Application.Windows.Item(1).Width = aw   Application.Windows.Item(1).Height = ah / 2   Application.Windows.Item(1).Left = 0   Application.Windows.Item(2).Width = aw   Application.Windows.Item(2).Height = ah / 2   Application.Windows.Item(2).Top = ah / 2   Application.Windows.Item(2).Left = 0 }  /*以下示例将指定表中第一列的宽度设置为 80 磅（每英寸 72 磅）。*/ Application.ActivePresentation.Slides.Item(2).Shapes.Item(5).Table.Columns.Item(1).Width = 80` |

适用环境：web

适用平台：windows/linux