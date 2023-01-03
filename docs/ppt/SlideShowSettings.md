**SlideShowSettings**



代表演示文稿的幻灯片放映设置。

**方法**

|                                                              | 名称    | 说明                                                         |
| ------------------------------------------------------------ | ------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Run** | 运行指定演示文稿的幻灯片放映。返回一个 **SlideShowWindow** 对象。 |

**属性**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AdvanceMode**       | 返回或设置一个值，该值指示幻灯片放映的切换方式。可读/写 **PpSlideShowAdvanceMode** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**       | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndingSlide**       | 返回或设置指定幻灯片放映中最后显示的幻灯片。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LoopUntilStopped**  | 决定指定幻灯片放映是否持续循环播放，直到用户按 Esc。**MsoTriState** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NamedSlideShows**   | 返回一个 **NamedSlideShows** 集合，该集合代表指定演示文稿中所有命名幻灯片放映（自定义幻灯片放映）。每个命名幻灯片放映或自定义幻灯片放映都是指定演示文稿的用户选择的子集。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**            | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PointerColor**      | 以 **ColorFormat** 对象的形式返回指定演示文稿的指针颜色。该颜色随演示文稿一起保存，并且是该演示文稿每次运行时的默认绘图笔颜色。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RangeType**         | 返回或设置运行幻灯片放映的类型。可读/写 **PpSlideShowRangeType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowScrollbar**     | 属性值为 **MsoTrue** 时，在以浏览模式放映幻灯片的过程中显示滚动条。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowType**          | 返回或设置指定幻灯片放映的放映类型。可读/写 **PpSlideShowType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowWithAnimation** | 决定指定的幻灯片放映是否显示已分配动画设置的形状。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowWithNarration** | 决定显示指定幻灯片放映时是否有旁白。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideShowName**     | 返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（**ActionSetting** 对象）；返回或设置自定义幻灯片放映的名称以便打印（**PrintOptions** 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（**PublishObject** 对象）。**String** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StartingSlide**     | 返回或设置指定幻灯片放映中要显示的第一张幻灯片。可读/写 **Long** 类型。 |

**成员方法**

#### **SlideShowSettings.Run**

运行指定演示文稿的幻灯片放映。返回一个 **SlideShowWindow** 对象。

**语法**

**express.Run()**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**返回值**

SlideShowWindow

**说明**

要运行自定义幻灯片放映，请将 **RangeType** 属性设置为 **ppShowNamedSlideShow**，并将 **SlideShowName** 属性设置为要运行的自定义放映的名称。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动演示文稿播放全屏幻灯片放映，并禁用快捷键。*/ function test(){ 　　　　ActivePresentation.SlideShowSettings.ShowType = ppShowSpeaker 　　　　ActivePresentation.SlideShowSettings.Run().View.AcceleratorsEnabled = false}` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
|                                                              |
| `/*本示例运行名为“Quick Show”的幻灯片放映。*/ function test(){ 　　　　ActivePresentation.SlideShowSettings.RangeType = ppShowNamedSlideShow 　　　　ActivePresentation.SlideShowSettings.SlideShowName = "Quick Show" 　　　　ActivePresentation.SlideShowSettings.Run() }` |

**成员属性**

#### **SlideShowSettings.AdvanceMode**

返回或设置一个值，该值指示幻灯片放映的切换方式。可读/写 **PpSlideShowAdvanceMode** 类型。

**语法**

**express.AdvanceMode**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**说明**

| PpSlideShowAdvanceMode 可以是下列 PpSlideShowAdvanceMode 常量之一。 |
| ------------------------------------------------------------ |
| ppSlideShowManualAdvance                                     |
| ppSlideShowRehearseNewTimings                                |
| ppSlideShowUseSlideTimings                                   |

#### **SlideShowSettings.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ 　　　　let shpOle = ActivePresentation.Slides.Item(1).Shapes        for(let i = 1; i <= shpOle.Count; i++) {    　　　　 if(shpOle.Item(i).Type == msoLinkedOLEObject) { 	　　　　MsgBox(shpOle.Item(i).OLEFormat.Application.Name)   　　　　  } 　　　　} }` |

#### **SlideShowSettings.EndingSlide**

返回或设置指定幻灯片放映中最后显示的幻灯片。可读/写 **Long** 类型。

**语法**

**express.EndingSlide**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例运行活动演示文稿的一个幻灯片放映，该放映从第二张幻灯片开始，到第四张幻灯片结束。*/ function test(){ 　　　　let SliSet = ActivePresentation.SlideShowSettings    　　　　 SliSet.RangeType = ppShowSlideRange   　　　　  SliSet.StartingSlide = 2   　　　　  SliSet.EndingSlide = 4    　　　　 SliSet.Run() }` |

#### **SlideShowSettings.LoopUntilStopped**

决定指定幻灯片放映是否持续循环播放，直到用户按 Esc。**MsoTriState** 类型，可读/写。

**语法**

**express.LoopUntilStopped**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。    |
| ------------------------------------------------ |
| msoCTrue                                         |
| msoFalse                                         |
| msoTriStateMixed                                 |
| msoTriStateToggle                                |
| msoTrue 指定幻灯片放映持续循环，直到用户按 Esc。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例开始当前演示文稿的一个幻灯片放映，该演示文稿自动切换幻灯片（按存储的时间间隔），并循环放映直到用户按 Esc。*/ function test(){ 　　　　let SliSet = ActivePresentation.SlideShowSettings   　　　　  SliSet.AdvanceMode = ppSlideShowUseSlideTimings   　　　　  SliSet.LoopUntilStopped = msoTrue    　　　　 SliSet.Run()}` |

#### **SlideShowSettings.NamedSlideShows**

返回一个 **NamedSlideShows** 集合，该集合代表指定演示文稿中所有命名幻灯片放映（自定义幻灯片放映）。每个命名幻灯片放映或自定义幻灯片放映都是指定演示文稿的用户选择的子集。只读。

**语法**

**express.NamedSlideShows**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**说明**

使用 **NamedSlideShows** 对象的 **Add** 方法创建命名幻灯片放映。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为当前演示文稿添加一个命名幻灯片放映“Quick Show”（该放映包含第二、第七和第九张幻灯片），然后运行该放映。*/ function test(){ 　　　　let qSlides = [] 　　　　let SliID = ActivePresentation.Slides    　　　　 qSlides[1] = SliID.Item(2).SlideID   　　　　  qSlides[2] = SliID.Item(7).SlideID     　　　　qSlides[3] = SliID.Item(9).SlideID  　　　　let SliSet = ActivePresentation.SlideShowSettings    　　　　 SliSet.RangeType = ppShowNamedSlideShow    　　　　 SliSet.NamedSlideShows.Add("Quick Show", qSlides)   　　　　  SliSet.SlideShowName = "Quick Show"    　　　　 SliSet.Run() }` |

#### **SlideShowSettings.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShapeOval,50,50,300,150).TextFrame   　　　　  myShapes.TextRange.Text = "Test text"   　　　　  myShapes.Parent.Rotation = 45 }` |

#### **SlideShowSettings.PointerColor**

以 **ColorFormat** 对象的形式返回指定演示文稿的指针颜色。该颜色随演示文稿一起保存，并且是该演示文稿每次运行时的默认绘图笔颜色。只读。

**语法**

**express.PointerColor**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**说明**

要将指针更改为绘图笔，可将 **PointerType** 属性设置为 **ppSlideShowPointerPen**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿的默认绘图笔颜色设为蓝色，然后开始一个幻灯片放映，将指针改变为绘图笔，并设定绘图笔颜色仅在此幻灯片放映中为红色。*/ function test(){ 　　　　let SliSet = ActivePresentation.SlideShowSettings  　　　　   SliSet.PointerColor.RGB = (0, 0, 255)          //blue   　　　　  SliSet.Run().View.PointerColor.RGB = (255, 0, 0)      //red 　　　　    SliSet.Run().View.PointerType = ppSlideShowPointerPen }` |

#### **SlideShowSettings.RangeType**

返回或设置运行幻灯片放映的类型。可读/写 **PpSlideShowRangeType** 类型。

**语法**

**express.RangeType**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**说明**

| PpSlideShowRangeType 可以是下列 PpSlideShowRangeType 常量之一。 |
| ------------------------------------------------------------ |
| ppShowAll                                                    |
| ppShowNamedSlideShow                                         |
| ppShowSlideRange                                             |

若要打印在 **PrintRanges** 集合中定义的幻灯片范围，必须先将 **RangeType** 属性设置为 **ppPrintSlideRange**。将 **RangeType** 设置为除 **ppPrintSlideRange** 之外的值会使在 **PrintRanges** 集合中定义的范围失效。但是这不会以任何方式影响 **PrintRanges** 集合的内容。也就是说，如果定义了一些打印范围，将 **RangeType** 属性设置为除 **ppPrintSlideRange** 之外的值，然后将 **RangeType** 重新设置为 **ppPrintSlideRange**，则之前定义的打印范围将保持不变。

指定 **PrintOut** 方法的 *To* 和 *From* 参数将设置该属性的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例运行名为“Quick Show”的幻灯片放映。*/ function test(){ 　　　　ActivePresentation.SlideShowSettings.RangeType = ppShowNamedSlideShow 　　　　ActivePresentation.SlideShowSettings.SlideShowName = "Quick Show" 　　　　ActivePresentation.SlideShowSettings.Run() }` |

#### **SlideShowSettings.ShowScrollbar**

属性值为 **MsoTrue** 时，在以浏览模式放映幻灯片的过程中显示滚动条。可读/写 **MsoTriState** 类型。

**语法**

**express.ShowScrollbar**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**说明**

在设置 **ShowScrollbar** 属性之前使用 **ShowType** 属性。

------

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue 不用于此属性。                       |
| msoFalse                                      |
| msoTriStateMixed 不用于此属性。               |
| msoTriStateToggle 不用于此属性                |
| msoTrue                                       |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定在某个窗口中显示当前演示文稿的幻灯片放映，并且在幻灯片放映时显示用于浏览幻灯片的滚动条。*/ function ShowSlideShowScrollBar() {     ActivePresentation.SlideShowSettings.ShowType = ppShowTypeWindow     ActivePresentation.SlideShowSettings.ShowScrollBar = msoTrue }` |

#### **SlideShowSettings.ShowType**

返回或设置指定幻灯片放映的放映类型。可读/写 **PpSlideShowType** 类型。

**语法**

**express.ShowType**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**说明**

| PpSlideShowType 可以是下列 PpSlideShowType 常量之一。 |
| ----------------------------------------------------- |
| ppShowTypeKiosk                                       |
| ppShowTypeSpeaker                                     |
| ppShowTypeWindow                                      |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在窗口中运行当前演示文稿的一个幻灯片放映，从第二张幻灯片开始，到第四张幻灯片结束。新的幻灯片放映窗口位于屏幕左上角，宽度和高度均为 300 磅。*/ function test(){ 　　　　let SliSet = ActivePresentation.SlideShowSettings    　　　　 SliSet.RangeType = ppShowSlideRange    　　　　 SliSet.StartingSlide = 2   　　　　  SliSet.EndingSlide = 4  　　　　   SliSet.ShowType = ppShowTypeWindow 　　　　let SliSetR = SliSet.Run()   　　　　  SliSetR.Left = 0 　　　　    SliSetR.Top = 0   　　　　  SliSetR.Width = 300    　　　　 SliSetR.Height = 300 }` |

#### **SlideShowSettings.ShowWithAnimation**

决定指定的幻灯片放映是否显示已分配动画设置的形状。可读/写 **MsoTriState** 类型。

**语法**

**express.ShowWithAnimation**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。      |
| -------------------------------------------------- |
| msoCTrue                                           |
| msoFalse                                           |
| msoTriStateMixed                                   |
| msoTriStateToggle                                  |
| msoTrue 指定的幻灯片放映显示已分配动画设置的形状。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例运行活动演示文稿的幻灯片放映，关闭动画和旁白。*/ function test(){ 　　　　ActivePresentation.SlideShowSettings.ShowWithAnimation = msoFalse 　　　　ActivePresentation.SlideShowSettings.ShowWithNarration = msoFalse 　　　　ActivePresentation.SlideShowSettings.Run() }` |

#### **SlideShowSettings.ShowWithNarration**

决定显示指定幻灯片放映时是否有旁白。可读/写 **MsoTriState** 类型。

**语法**

**express.ShowWithNarration**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue 显示指定幻灯片放映时有旁白。          |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例运行活动演示文稿的幻灯片放映，关闭动画和旁白。*/ function test(){ 　　　　ActivePresentation.SlideShowSettings.ShowWithAnimation = msoFalse 　　　　ActivePresentation.SlideShowSettings.ShowWithNarration = msoFalse 　　　　ActivePresentation.SlideShowSettings.Run() }` |

#### **SlideShowSettings.SlideShowName**

返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（**ActionSetting** 对象）；返回或设置自定义幻灯片放映的名称以便打印（**PrintOptions** 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（**PublishObject** 对象）。**String** 类型，可读/写。

**语法**

**express.SlideShowName**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**说明**

必须将 **RangeType** 属性设置为 **ppPrintNamedSlideShow** 才能打印自定义幻灯片放映。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印名为“tech talk”的现有自定义幻灯片放映。*/ function test(){ 　　　　ActivePresentation.PrintOptions.RangeType = ppPrintNamedSlideShow 　　　　ActivePresentation.PrintOptions.SlideShowName = "tech talk" 　　　　ActivePresentation.PrintOut() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例将当前演示文稿保存为 HTML 4.0 文件格式，文件名为“mallard.htm”。然后显示一条消息，说明正将当前名称的演示文稿保存为WPP 和 HTML 两种格式。*/ function test(){ 　　　　let ActivePub = ActivePresentation.PublishObjects.Item(1) 　　　　let PresName = ActivePub.SlideShowName 　　　　    ActivePub.SourceType = ppPublishAll 　　　　    ActivePub.FileName = "C:\\HTMLPres\\mallard.htm" 　　　　    ActivePub.HTMLVersion = ppHTMLv4 　　　　    MsgBox("Saving presentation " + "'" + PresName + "'" + " inWPP" + "\n" + "\r" + " format and HTML version 4.0 format") 　　　　    ActivePub.Publish }` |

#### **SlideShowSettings.StartingSlide**

返回或设置指定幻灯片放映中要显示的第一张幻灯片。可读/写 **Long** 类型。

**语法**

**express.StartingSlide**

*express*   一个代表 **SlideShowSettings** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例运行活动演示文稿的一个幻灯片放映，该放映从第二张幻灯片开始，到第四张幻灯片结束。*/ function test(){ 　　　　ActivePresentation.SlideShowSettings.RangeType = ppShowSlideRange 　　　　ActivePresentation.SlideShowSettings.StartingSlide = 2 　　　　ActivePresentation.SlideShowSettings.EndingSlide = 4 　　　　ActivePresentation.SlideShowSettings.Run() }` |

适用环境：web

适用平台：windows/linux