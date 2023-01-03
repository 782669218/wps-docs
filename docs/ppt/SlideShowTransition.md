**SlideShowTransition**



包含关于幻灯片放映中指定幻灯片的切换方式的信息。

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AdvanceOnClick**     | 决定在幻灯片放映期间单击指定的幻灯片后是否会切换。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AdvanceOnTime**      | 决定指定的幻灯片在经过指定时间后是否自动切换。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AdvanceTime**        | 返回或设置以秒为单位的时间长度，该段时间过后，指定的幻灯片将会切换。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**        | 返回一个 **Application** 对象，该对象表示指定对象的创建者    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EntryEffect**        | 对于 **AnimationSettings** 对象，此属性返回或设置应用于指定形状中动画的特殊效果。对于 **SlideShowTransition** 对象，此属性返回或设置应用于指定幻灯片切换的特殊效果。PpEntryEffect 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hidden**             | 决定幻灯片放映过程中指定的幻灯片是否隐藏。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LoopSoundUntilNext** | 指定是否循环播放为指定的幻灯片切换所设置的声音，直到下一个声音开始。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**             | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SoundEffect**        | 返回一个 **SoundEffect** 对象，该对象代表切换到指定幻灯片时播放的声音。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Speed**              | 返回或设置一个 **PpTransitionSpeed** 常量，该常量代表切换到指定幻灯片的速度。可读/写。 |

**成员属性**

#### **SlideShowTransition.AdvanceOnClick**

决定在幻灯片放映期间单击指定的幻灯片后是否会切换。可读/写 **MsoTriState** 类型。

**语法**

**express.AdvanceOnClick**

*express*   一个代表 **SlideShowTransition** 对象的变量。

**说明**

要设置幻灯片在某段时间后自动切换，请将 **AdvanceOnTime** 属性设置为 **True**，并将 **AdvanceTime** 属性设置为希望幻灯片放映的时间。如果同时将 **AdvanceOnClick** 和 **AdvanceOnTime** 属性设置为 **True**，则幻灯片将在被单击时或经过指定的时间后（无论哪个先发生）进行切换。

| MsoTriState 可以是下列 MsoTriState 常量之一。            |
| -------------------------------------------------------- |
| msoCTrue                                                 |
| msoFalse                                                 |
| msoTriStateMixed                                         |
| msoTriStateToggle                                        |
| msoTrue 在幻灯片放映过程中，单击指定的幻灯片后进行切换。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动演示文稿中的第一张幻灯片在五秒钟过去后前进，或在单击鼠标时前进（具体取决于哪个事件首先发生）。*/ function test(){ 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnClick = msoTrue 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnTime = msoTrue 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceTime = 5 }` |

#### **SlideShowTransition.AdvanceOnTime**

决定指定的幻灯片在经过指定时间后是否自动切换。可读/写 **MsoTriState** 类型。

**语法**

**express.AdvanceOnTime**

*express*   一个代表 **SlideShowTransition** 对象的变量。

**说明**

使用 **AdvanceTime** 属性可以指定幻灯片自动切换前需等待的秒数。将 **SlideShowSettings** 对象的 **AdvanceMode** 属性设置为 **ppSlideShowUseSlideTimings** 可使幻灯片的间隔设置对整个幻灯片放映有效。

| MsoTriState 可以是下列 MsoTriState 常量之一。  |
| ---------------------------------------------- |
| msoCTrue                                       |
| msoFalse                                       |
| msoTriStateMixed                               |
| msoTriStateToggle                              |
| msoTrue 指定的幻灯片在经过指定时间后自动切换。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动演示文稿中的第一张幻灯片在五秒钟过去后前进，或在单击鼠标时前进（具体取决于哪个事件首先发生）。*/ function test(){ 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnClick = msoTrue 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnTime = msoTrue 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceTime = 5 }` |

#### **SlideShowTransition.AdvanceTime**

返回或设置以秒为单位的时间长度，该段时间过后，指定的幻灯片将会切换。可读/写。

**语法**

**express.AdvanceTime**

*express*   一个代表 **SlideShowTransition** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动演示文稿中的第一张幻灯片在五秒钟过去后前进，或在单击鼠标时前进（具体取决于哪个事件首先发生）。*/ function test(){ 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnClick = msoTrue 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceOnTime = msoTrue 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.AdvanceTime = 5 }` |

#### **SlideShowTransition.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者

**语法**

**express.Application**

*express*   一个代表 **SlideShowTransition** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ 　　　　let shpOle = ActivePresentation.Slides.Item(1).Shapes        for(let i = 1; i <= shpOle.Count; i++) {    　　　　 if(shpOle.Item(i).Type == msoLinkedOLEObject) {    　　　　     MsgBox(shpOle.Item(i).OLEFormat.Application.Name)  　　　　   } 　　　　}}` |

#### **SlideShowTransition.EntryEffect**

对于 **AnimationSettings** 对象，此属性返回或设置应用于指定形状中动画的特殊效果。对于 **SlideShowTransition** 对象，此属性返回或设置应用于指定幻灯片切换的特殊效果。PpEntryEffect 类型，可读写。

**语法**

**express.EntryEffect**

*express*   一个代表 **SlideShowTransition** 对象的变量。

**说明**

如果将指定形状的 **TextLevelEffect** 属性设为 **ppAnimateLevelNone**（默认值）或将 **Animate** 属性设为 **False**，您将看不到通过 **EntryEffect** 属性应用的特殊效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前演示文稿中添加一张标题幻灯片，并设置该标题在幻灯片放映中动画时从右侧飞入。*/ function test(){ 　　　　let PreTit = ActivePresentation.Slides.Add(1, ppLayoutTitleOnly).Shapes.Item(1)    　　　　 PreTit.TextFrame.TextRange.Text = "Sample title"    　　　　 PreTit.AnimationSettings.TextLevelEffect = ppAnimateByAllLevels    　　　　 PreTit.AnimationSettings.EntryEffect = ppEffectFlyFromRight    　　　　 PreTit.AnimationSettings.Animate = true }` |

#### **SlideShowTransition.Hidden**

决定幻灯片放映过程中指定的幻灯片是否隐藏。可读/写 **MsoTriState** 类型。

**语法**

**express.Hidden**

*express*   一个代表 **SlideShowTransition** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue 幻灯片放映过程中隐藏指定的幻灯片。    |

**示例**

本示例将活动演示文稿中的第二张幻灯片设为隐藏。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `ActivePresentation.Slides.Item(2).SlideShowTransition.Hidden = msoTrue` |

#### **SlideShowTransition.LoopSoundUntilNext**

指定是否循环播放为指定的幻灯片切换所设置的声音，直到下一个声音开始。可读/写 **MsoTriState** 类型。

**语法**

**express.LoopSoundUntilNext**

*express*   一个代表 **SlideShowTransition** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| msoCTrue                                                     |
| msoFalse                                                     |
| msoTriStateMixed                                             |
| msoTriStateToggle                                            |
| msoTrue 循环播放为指定的幻灯片切换所设置的声音，直到下一个声音开始。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定转换到当前演示文稿第二张幻灯片时开始播放文件 Dudududu.wav，并持续播放直到开始下一个声音文件。*/ function test(){ 　　　　ActivePresentation.Slides.Item(2).SlideShowTransition.SoundEffect.ImportFromFile("C:\\sndsys\\dudududu.wav") 　　　　ActivePresentation.Slides.Item(2).SlideShowTransition.LoopSoundUntilNext = msoTrue }` |

#### **SlideShowTransition.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **SlideShowTransition** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame 　　　　myShapes.TextRange.Text = "Test text" 　　　　myShapes.Parent.Rotation = 45 }` |

#### **SlideShowTransition.SoundEffect**

返回一个 **SoundEffect** 对象，该对象代表切换到指定幻灯片时播放的声音。

**语法**

**express.SoundEffect**

*express*   一个代表 **SlideShowTransition** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置在当前演示文稿第一张幻灯片的第一个形状演示动画时播放文件 Bass.wav。*/ function test(){ 　　　　let ShapeSet = ActivePresentation.Slides.Item(1).Shapes.Item(1).AnimationSettings 　　　　    ShapeSet.Animate = true 　　　　    ShapeSet.TextLevelEffect = ppAnimateByAllLevels 　　　　    ShapeSet.SoundEffect.ImportFromFile("C:\\bass.wav") }` |

#### **SlideShowTransition.Speed**

返回或设置一个 **PpTransitionSpeed** 常量，该常量代表切换到指定幻灯片的速度。可读/写。

**语法**

**express.Speed**

*express*   一个代表 **SlideShowTransition** 对象的变量。

**说明**

| PpTransitionSpeed 可以是下列 PpTransitionSpeed 常量之一。 |
| --------------------------------------------------------- |
| ppTransitionSpeedFast                                     |
| ppTransitionSpeedMedium                                   |
| ppTransitionSpeedMixed                                    |
| ppTransitionSpeedSlow                                     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置切换到当前演示文稿第一张幻灯片的特殊效果，并指定切换为快速。*/ function test(){ 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.EntryEffect = ppEffectStripsDownLeft 　　　　ActivePresentation.Slides.Item(1).SlideShowTransition.Speed = ppTransitionSpeedFast }` |

适用环境：web

适用平台：windows/linux