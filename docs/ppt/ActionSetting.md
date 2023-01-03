**ActionSetting**



包含指定形状或文本范围在幻灯片放映中对鼠标动作的反应的有关信息。

**说明**

**ActionSetting** 对象是 **ActionSettings** 集合的成员。

**ActionSettings** 集合包含一个代表在幻灯片放映中用户单击指定对象时该对象的反应的 **ActionSetting** 对象，以及一个代表在幻灯片放映中用户将鼠标指针移动到指定对象上时该对象的反应的 **ActionSetting** 对象。

如果设置了 **ActionSetting** 对象的属性但没有效果，请确认已将 **Action** 属性设为适当的值。

**属性**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Action**        | 返回或设置在幻灯片放映过程中单击指定形状或将鼠标指针定位到形状上时将发生的动作的类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActionVerb**    | 返回或设置包含 OLE 动作的字符串，该 OLE 动作将在用户在幻灯片放映过程中单击指定形状或将鼠标指针移过该形状时运行。要使该属性可影响幻灯片放映操作，必须先将 Action 属性设置为 ppActionOLEVerb。String 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AnimateAction** | 发生指定的鼠标操作时，如果指定形状的颜色立即反转，则为 MsoTrue。MsoTriState 类型，可读/写 。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**   | 返回一个 Application 对象，该对象表示指定对象的创建者。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyperlink**     | 返回一个 Hyperlink 对象，该对象代表指定形状的超链接。要在幻灯片放映过程中激活超链接，必须将 Action 属性设为 ppActionHyperlink。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**        | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Run**           | 返回或设置演示文稿或宏的名称，当在幻灯片放映过程中单击指定形状或鼠标指针经过该形状时，将运行该演示文稿或宏。Action 属性必须设置为 ppActionRunMacro 或 ppActionRunProgram，此属性才能影响幻灯片的放映动作。可读/写。String 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAndReturn** | 决定 WPP是否返回初始的幻灯片放映以及返回的条件。可读/写。MsoTriState 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideShowName** | 返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（ActionSetting 对象）；返回或设置自定义幻灯片放映的名称以便打印（PrintOptions 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（PublishObject 对象）。String 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SoundEffect**   | 返回一个 SoundEffect 对象，该对象代表切换到指定幻灯片时播放的声音。 |

**成员属性**

#### **ActionSetting.Action**

返回或设置在幻灯片放映过程中单击指定形状或将鼠标指针定位到形状上时将发生的动作的类型。

**语法**

**express.Action**

*express*   一个代表 **ActionSetting** 对象的变量。

**说明**

可以是下列 **PpActionType** 常量之一。

| ppActionEndShow         |
| ----------------------- |
| ppActionFirstSlide      |
| ppActionHyperlink       |
| ppActionLastSlide       |
| ppActionLastSlideViewed |
| ppActionMixed           |
| ppActionNamedSlideShow  |
| ppActionNextSlide       |
| ppActionNone            |
| ppActionOLEVerb         |
| ppActionPlay            |
| ppActionPreviousSlide   |
| ppActionRunMacro        |
| ppActionRunProgram      |

可将 **Action** 属性与 **ActionSetting** 对象的其他属性结合使用，如下表所示。

| 如果将Action 属性设置为此值 | 使用此属性    | 进行此操作                                                   |
| --------------------------- | ------------- | ------------------------------------------------------------ |
| ppActionHyperlink           | Hyperlink     | 设置超链接的属性，在幻灯片放映过程中为响应形状上的鼠标动作将转至该超链接。 |
| ppActionRunProgram          | Run           | 返回或设置在幻灯片放映过程中为响应形状上的鼠标动作而运行的程序的名称。 |
| ppActionRunMacro            | Run           | 返回或设置在幻灯片放映过程中为响应形状上的鼠标动作而运行的宏的名称。 |
| ppActionOLEVerb             | ActionVerb    | 设置在幻灯片放映过程中为响应形状上的鼠标动作而调用的 OLE 动作。 |
| ppActionNamedSlideShow      | SlideShowName | 设置自定义放映名称以响应幻灯片放映过程中对形状的鼠标操作。   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动演示文稿第一张幻灯片的第三个形状（OLE 对象），使其在幻灯片放映过程中当鼠标移过它时播放。可读/写 Long 类型。*/ function test(){ let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).ActionSettings.Item(ppMouseOver) Act.ActionVerb = "Play" Act.Action = ppActionOLEVerb }` |

#### **ActionSetting.ActionVerb**

返回或设置包含 OLE 动作的字符串，该 OLE 动作将在用户在幻灯片放映过程中单击指定形状或将鼠标指针移过该形状时运行。要使该属性可影响幻灯片放映操作，必须先将 Action 属性设置为 ppActionOLEVerb。String 类型，可读/写。

**语法**

**express.ActionVerb**

*express*   一个代表 **ActionSetting** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将第一张幻灯片的第三个形状设置为：只要鼠标指针在幻灯片放映时移过该形状就会被播放。但第三个形状必须是支持“Play”动词的 OLE 对象。*/ function test(){ 　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).ActionSettings.Item(ppMouseOver) 　　　　Act.ActionVerb = "Play" 　　　　Act.Action = ppActionOLEVerb}` |

#### **ActionSetting.AnimateAction**

发生指定的鼠标操作时，如果指定形状的颜色立即反转，则为 MsoTrue。MsoTriState 类型，可读/写 。

**语法**

**express.AnimateAction**

*express*   一个代表 **ActionSetting** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue                                       |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置在幻灯片放映过程中，单击活动演示文稿第一张幻灯片的第三个形状时播放掌声，并立即反转该形状的颜色。 */ function test(){ 　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).ActionSettings.Item(ppMouseClick) 　　　　Act.SoundEffect.Name = "applause" 　　　　Act.AnimateAction = msoTrue }` |

#### **ActionSetting.Application**

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **ActionSetting** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。 */ function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test(){ 　　　　for(let i= 1;i< = ActivePresentation.Slides.Item(1).Shapes.Count;i++) {     　　　　if(ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {         　　　　MsgBox(ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)     　　　　} 　　　　} }` |

#### **ActionSetting.Hyperlink**

返回一个 Hyperlink 对象，该对象代表指定形状的超链接。要在幻灯片放映过程中激活超链接，必须将 Action 属性设为 ppActionHyperlink。只读。

**语法**

**express.Hyperlink**

*express*   一个代表 **ActionSetting** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对当前演示文稿中第一张幻灯片上的第一个形状进行设置，使其在幻灯片演示过程中，如果被单击则跳转到 WPS 网站上。*/ function test(){ 　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).ActionSettings.Item(ppMouseClick) 　　　　Act.Action = ppActionHyperlink 　　　　Act.Hyperlink.Address = "http://www.wps.cn/" }` |

#### **ActionSetting.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ActionSetting** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 */ function test(){ 　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes 　　　　let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame 　　　　addShp.TextRange.Text = "Test text" 　　　　addShp.Parent.Rotation = 45 }` |

#### **ActionSetting.Run**

返回或设置演示文稿或宏的名称，当在幻灯片放映过程中单击指定形状或鼠标指针经过该形状时，将运行该演示文稿或宏。Action 属性必须设置为 ppActionRunMacro 或 ppActionRunProgram，此属性才能影响幻灯片的放映动作。可读/写。String 类型。

**语法**

**express.Run**

*express*   一个代表 **ActionSetting** 对象的变量。

**说明**

如果 **Action** 属性的值为 **ppActionRunMacro**，则指定的字符串值应为一个当前已加载的全局宏的名称。如果 **Action** 属性的值为 **ppActionRunProgram**，则指定的字符串应为一个程序的完整路径和文件名。

可以将 **Run** 属性设为无参数的或单个 *Shape* 或 *Object* 参数的宏。幻灯片放映时被单击的形状将被作为此参数传送。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定在幻灯片放映过程中的任何时候，如果鼠标指针经过形状，将运行宏 CalculateTotal。*/ function test(){ 　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).ActionSettings.Item(ppMouseOver) 　　　　Act.Action = ppActionRunMacro 　　　　Act.Run = "CalculateTotal" 　　　　Act.AnimateAction = true }` |

#### **ActionSetting.ShowAndReturn**

决定 WPP是否返回初始的幻灯片放映以及返回的条件。可读/写。MsoTriState 类型。

**语法**

**express.ShowAndReturn**

*express*   一个代表 **ActionSetting** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| msoCTrue                                                     |
| msoFalse 默认值。WPP不会从未激活的自定义幻灯片放映返回初始幻灯片放映。 |
| msoTriStateMixed                                             |
| msoTriStateToggle                                            |
| msoTrue：WPP从取消激活的自定义幻灯片放映返回初始幻灯片放映，该自定义幻灯片放映通过初始演示文稿的 Hyperlink 或 ActionSetting 对象激活。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置当前演示文稿中第一张幻灯片上第五个形状的鼠标单击动作为：显示名为“techtalk”的自定义幻灯片放映。当该自定义幻灯片放映结束后，将自动返回演示文稿在鼠标单击前的最初状态。*/ function test(){ 　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(5).ActionSettings.Item(ppMouseClick) 　　　　Act.Action = ppActionNamedSlideShow 　　　　Act.SlideShowName = "techtalk" 　　　　Act.ShowandReturn = msoTrue }` |

#### **ActionSetting.SlideShowName**

返回或设置自定义幻灯片放映的名称，该幻灯片放映将运行以响应幻灯片放映期间在形状上的鼠标动作（ActionSetting 对象）；返回或设置自定义幻灯片放映的名称以便打印（PrintOptions 对象）；返回或设置发布为 Web 演示文稿的自定义幻灯片放映的名称（PublishObject 对象）。String 类型，可读/写。

**语法**

**express.SlideShowName**

*express*   一个代表 **ActionSetting** 对象的变量。

**说明**

必须将 **RangeType** 属性设置为 **ppPrintNamedSlideShow** 才能打印自定义幻灯片放映。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印名为“tech talk”的现有自定义幻灯片放映。*/ function test(){ 　　　　let Pri = Application.ActivePresentation.PrintOptions 　　　　Pri.RangeType = ppPrintNamedSlideShow 　　　　Pri.SlideShowName = "tech talk" 　　　　ActivePresentation.PrintOut() } ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例将当前演示文稿保存为 HTML 4.0 文件格式，文件名为“mallard.htm”。然后显示一条消息，说明正将当前名称的演示文稿保存为WPP 和 HTML 两种格式。*/ function test(){ 　　　　let Pub = ActivePresentation.PublishObjects.Item(1) 　　　　let PresName = Pub.SlideShowName 　　　　Pub.SourceType = ppPublishAll 　　　　Pub.FileName = "C:\\HTMLPres\\mallard.htm" 　　　　Pub.HTMLVersion = ppHTMLVersion4 　　　　MsgBox("Saving presentation " + "'" + PresName + "'" + " inWPP" + "\n" + "\r" + " format and HTML version 4.0 format") 　　　　Pub.Publish() }` |

#### **ActionSetting.SoundEffect**

返回一个 SoundEffect 对象，该对象代表切换到指定幻灯片时播放的声音。

**语法**

**express.SoundEffect**

*express*   一个代表 **ActionSetting** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置在当前演示文稿第一张幻灯片的第一个形状演示动画时播放文件 Bass.wav。*/ function test(){ 　　　　let Ani = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).AnimationSettings 　　　　Ani.Animate = true 　　　　Ani.TextLevelEffect = ppAnimateByAllLevels 　　　　Ani.SoundEffect.ImportFromFile("c:\\bass.wav") }` |

适用环境：web

适用平台：windows/linux