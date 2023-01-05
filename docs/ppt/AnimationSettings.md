**AnimationSettings**



代表幻灯片放映时应用于指定形状的动画的特殊效果。

**属性**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AdvanceMode**          | 返回或设置一个值，该值指示指定形状的动画是仅在被单击时换页还是在经过指定时间后自动换页。可读/写 **PpAdvanceMode** 类型。如果该形状不会有动画效果，请确保将 **TextLevelEffect** 属性设置为除 **ppAnimateLevelNone** 之外的值，并且将 **Animate** 属性设置为 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AdvanceTime**          | 返回或设置以秒为单位的时间，在该时间过后指定的形状将开始播放动画。可读/写 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AfterEffect**          | 返回或设置 **PpAfterEffect** 常量，该常量指示指定的形状在生成之后是变暗、隐藏还是不更改。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Animate**              | 决定在幻灯片放映过程中指定的形状是否被赋予动画效果。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AnimateBackground**    | 如果指定的对象是自选图形，若该形状独立于其所包含的文本单独显示动画，则值为 **msoTrue**。如果指定的形状是图形对象，若指定的图形对象的背景（坐标轴和网格线）被赋予动画效果，则值为 **msoTrue**。仅适用于图形对象或带有经多个步骤才能显示的文本的自选图形。**MsoTriState** 类型，可读/写 。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AnimateTextInReverse** | 决定指定形状是否按相反顺序构造。仅应用于具有多个创建步骤的形状（例如，包含列表的形状）。可读/写。**MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AnimationOrder**       | 返回或设置一个整数，该整数代表指定形状在动画形状集合中的位置。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**          | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChartUnitEffect**      | 返回或设置一个值，该值指示图形区域是按序列、类别还是按元素播放动画。可读/写 **PpChartUnitEffect** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DimColor**             | 返回或设置一个 **ColorFormat** 对象，该对象代表指定形状生成后的颜色。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EntryEffect**          | 对于 **AnimationSettings** 对象，此属性返回或设置应用于指定形状中动画的特殊效果。对于 **SlideShowTransition** 对象，此属性返回或设置应用于指定幻灯片切换的特殊效果。**PpEntryEffect** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**               | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PlaySettings**         | 返回一个 **PlaySettings** 对象，该对象包含指定的媒体剪辑在幻灯片放映中的播放方式的信息。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SoundEffect**          | 返回一个 **SoundEffect** 对象，该对象代表切换到指定幻灯片时播放的声音。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextLevelEffect**      | 设置或返回一个 **PpTextLevelEffect** 常量，该常量表示指定形状中的文本是否由第一级段落、第二级段落或其他级段落（最高为第五级段落）激活动画显示。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextUnitEffect**       | 返回或设置一个 **PpTextUnitEffect** 常量，该常量指示指定形状中的文本是按段落、按字或是按字母被赋予动画效果。可读/写。 |

**成员属性**

#### **AnimationSettings.AdvanceMode**

返回或设置一个值，该值指示指定形状的动画是仅在被单击时换页还是在经过指定时间后自动换页。可读/写 **PpAdvanceMode** 类型。如果该形状不会有动画效果，请确保将 **TextLevelEffect** 属性设置为除 **ppAnimateLevelNone** 之外的值，并且将 **Animate** 属性设置为 **True**。

**语法**

**express.AdvanceMode**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

| PpAdvanceMode 可以是下列 PpAdvanceMode 常量之一。 |
| ------------------------------------------------- |
| ppAdvanceModeMixed                                |
| ppAdvanceOnClick                                  |
| ppAdvanceOnTime                                   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿第一张幻灯片的第二个形状设为五秒后自动播放动画。*/ function test(){ 　　　　let Ani = Appliaction.ActivePresentation.Slides.Item(1).Shapes.Item(2).AnimationSettings 　　　　Ani.AdvanceMode = ppAdvanceOnTime 　　　　Ani.AdvanceTime = 5 　　　　Ani.TextLevelEffect = ppAnimateByAllLevels 　　　　Ani.Animate = true }` |

#### **AnimationSettings.AdvanceTime**

返回或设置以秒为单位的时间，在该时间过后指定的形状将开始播放动画。可读/写

**语法**

**express.AdvanceTime**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

在指定的时间过后，指定的幻灯片动画不会自动开始，除非将动画的 AdvanceMode 属性设置为 ppAdvanceOnTime。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿第一张幻灯片上的第二个形状设为五秒后自动播放动画。 */ function test(){ 　　　　let Ani = Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).AnimationSettings 　　　　Ani.AdvanceMode = ppAdvanceOnTime 　　　　Ani.AdvanceTime = 5 　　　　Ani.TextLevelEffect = ppAnimateByAllLevels 　　　　Ani.Animate = true }` |

#### **AnimationSettings.AfterEffect**

返回或设置 **PpAfterEffect** 常量，该常量指示指定的形状在生成之后是变暗、隐藏还是不更改。可读/写。

**语法**

**express.AfterEffect**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

除非带有动画播放后效果的形状播放动画，且幻灯片上至少有另一个形状在其播放后播放动画；否则无法看到为该形状设置的动画播放后效果。要给一个形状赋予动画效果，必须将 **AnimationSettings** 对象的 **TextLevelEffect** 属性设置为 **ppAnimateLevelNone**以外的值，并且 **EntryEffect** 属性设置为除 **ppEffectNone** 以外的常量。另外，还必须将 **Animate** 属性必须设置为 **True**。若要更改幻灯片上形状的创建顺序，请使用 **AnimationOrder** 属性。

| PpAfterEffect 可以是下列 PpAfterEffect 常量之一。 |
| ------------------------------------------------- |
| ppAfterEffectDim                                  |
| ppAfterEffectHide                                 |
| ppAfterEffectHideOnClick                          |
| ppAfterEffectMixed                                |
| ppAfterEffectNothing                              |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定当前演示文稿中第一张幻灯片的标题在创建后变暗。但是，如果它是第一张幻灯片中最后的或唯一的形状，文本将不会变暗。*/ function test(){ 　　　　let Ani = Application.ActivePresentation.Slides.Item(1).Shapes.Title.AnimationSettings 　　　　Ani.Animate = true 　　　　Ani.TextLevelEffect = ppAnimateByAllLevels 　　　　Ani.AfterEffect = ppAfterEffectDim }` |

#### **AnimationSettings.Animate**

决定在幻灯片放映过程中指定的形状是否被赋予动画效果。可读/写 **MsoTriState** 类型。

**语法**

**express.Animate**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

要为一个形状赋予动画效果，则该形状的 **AnimationSettings** 对象的 **TextLevelEffect** 属性必须设置为除 **ppAnimateLevelNone** 之外的值，并且必须将 **Animate** 属性设置为 **True**，或将 **EntryEffect** 属性设置为除 **ppEffectNone** 之外的常量。

| MsoTriState 可以是下列 MsoTriState 常量之一。        |
| ---------------------------------------------------- |
| msoCTrue                                             |
| msoFalse                                             |
| msoTriStateMixed                                     |
| msoTriStateToggle                                    |
| msoTrue 在幻灯片放映过程中指定的形状被赋予动画效果。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定活动演示文稿第二张幻灯片的标题在创建后变暗。如果该标题是第二张幻灯片中的最后一个或唯一创建的形状，则该文本不变暗。*/ function test(){ 　　　　let Ani = Application.ActivePresentation.Slides.Item(2).Shapes.Title.AnimationSettings 　　　　Ani.TextLevelEffect = ppAnimateByAllLevels 　　　　Ani.AfterEffect = ppAfterEffectDim 　　　　Ani.Animate = msoTrue }` |

#### **AnimationSettings.AnimateBackground**

如果指定的对象是自选图形，若该形状独立于其所包含的文本单独显示动画，则值为 **msoTrue**。如果指定的形状是图形对象，若指定的图形对象的背景（坐标轴和网格线）被赋予动画效果，则值为 **msoTrue**。仅适用于图形对象或带有经多个步骤才能显示的文本的自选图形。**MsoTriState** 类型，可读/写 。

**语法**

**express.AnimateBackground**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

使用 **TextLevelEffect** 和 **TextUnitEffect** 属性可以控制附加到指定形状的文本的动画效果。

如果将该属性设置为 **MsoTrue** 并将 **TextLevelEffect** 属性设置为 **ppAnimateByAllLevels**，则形状及其文本将被同时赋予动画效果。如果将该属性设置为 **MsoTrue** 并将 **TextLevelEffect** 属性设置为除 **ppAnimateByAllLevels** 之外的值，则形状将在文本被赋予动画效果之前直接被赋予动画效果。

除非为指定形状赋予动画效果，否则看不到设置该属性的效果。要给一个形状赋予动画效果，必须将 **TextLevelEffect** 属性设置为 **ppAnimateLevelNone** 以外的值，并且将 **Animate** 属性设置为 **MsoTrue**，或者将 **EntryEffect** 属性设置为除 **ppEffectNone** 之外的常量。

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
| `/*本示例将创建一个包含文本的矩形。本示例还指定该形状应从右下角飞入，应从第一级段落显示文本，并且该形状应独立于其所包含的文本单独显示动画。在本示例中，EntryEffect 属性用于打开动画效果。*/ function AnimateTextBox() {     let addShp = Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShapeRectangle, 50, 200, 200, 200)         addShp.TextFrame.TextRange.Text = "Reason 1" + "\r" + "Reason 2" + "\r" + "Reason 3"     let Ani = addShp.AnimationSettings         Ani.EntryEffect = ppEffectFlyFromBottomRight         Ani.TextLevelEffect = ppAnimateByFirstLevel         Ani.TextUnitEffect = ppAnimateByParagraph         Ani.AnimateBackground = msoTrue }` |

#### **AnimationSettings.AnimateTextInReverse**

决定指定形状是否按相反顺序构造。仅应用于具有多个创建步骤的形状（例如，包含列表的形状）。可读/写。**MsoTriState** 类型。

**语法**

**express.AnimateTextInReverse**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue：指定形状按相反顺序构造。             |

指定形状必须被赋予动画效果，您才能看到设置该属性的效果。要给一个形状赋予动画效果，必须将 **AnimationSettings** 对象的 **TextLevelEffect** 属性设置为 **ppAnimateLevelNone** 以外的值，并且 **Animate** 属性必须设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在当前演示文稿的第一张幻灯片后添加一张幻灯片，然后设置标题文本，并为文本占位符添加三个项的列表，并设置该列表按相反顺序构造。*/ function test(){ 　　　　let myShape = Application.ActivePresentation.Slides.Add(2, ppLayoutText).Shapes    　　　　 myShape.Item(1).TextFrame.TextRange.Text = "Top Three Reasons" 　　　　let shape2 = myShape.Item(2)    　　　　 shape2.TextFrame.TextRange.Text = "Reason 1" + "\r" + "Reason 2" + "\r" + "Reason 3" 　　　　let Ani = shape2.AnimationSettings   　　　　  Ani.Animate = msoTrue   　　　　  Ani.TextLevelEffect = ppAnimateByFirstLevel   　　　　  Ani.AnimateTextInReverse = msoTrue }` |

#### **AnimationSettings.AnimationOrder**

返回或设置一个整数，该整数代表指定形状在动画形状集合中的位置。可读/写 **Long** 类型。

**语法**

**express.AnimationOrder**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

指定形状必须被赋予动画效果，您才能看到设置该属性的效果。要给一个形状赋予动画效果，必须将 **AnimationSettings** 对象的 **TextLevelEffect** 属性设置为 **ppAnimateLevelNone** 以外的值，并且 **Animate** 属性必须设置为 **True**。

**示例**

本示例指定活动演示文稿中第二张幻灯片的第二个形状在所有动画中第二个播放。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(2).Shapes.Item(2).AnimationSettings.AnimationOrder = 2` |

#### **AnimationSettings.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **AnimationSettings** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres)     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ 　　　　for(let i=1;i <= ActivePresentation.Slides.Item(1).Shapes.Count;i++) {    　　　　 if(ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {       　　　　  MsgBox(ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name) 　　　　    } 　　　　} }` |

#### **AnimationSettings.ChartUnitEffect**

返回或设置一个值，该值指示图形区域是按序列、类别还是按元素播放动画。可读/写 **PpChartUnitEffect** 类型。

**语法**

**express.ChartUnitEffect**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

如果图形未显示动画效果，请确保 **Animate** 属性已设置为 **True**。

| PpChartUnitEffect 可以是下列 PpChartUnitEffect 常量之一。 |
| --------------------------------------------------------- |
| ppAnimateByCategory                                       |
| ppAnimateByCategoryElements                               |
| ppAnimateBySeries                                         |
| ppAnimateBySeriesElements                                 |
| ppAnimateChartAllAtOnce                                   |
| ppAnimateChartMixed                                       |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动演示文稿第三张幻灯片的第二个形状按序列动画显示。为使本示例运行，第二个形状必须为一个图表。*/ function test(){ 　　　　let myShape = Application.ActivePresentation.Slides.Item(3).Shapes.Item(2) 　　　　let Ani = myShape.AnimationSettings  　　　　   Ani.ChartUnitEffect = ppAnimateBySeries 　　　　    Ani.EntryEffect = ppEffectFlyFromLeft  　　　　   Ani.Animate = true }` |

#### **AnimationSettings.DimColor**

返回或设置一个 **ColorFormat** 对象，该对象代表指定形状生成后的颜色。只读。

**语法**

**express.DimColor**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

如果没有达到预期的效果，请检查其他的创建设置。必须将 **AnimationSettings** 对象的 **TextLevelEffect** 属性设为非 **ppAnimateLevelNone** 的值，**AfterEffect** 属性设为 **ppAfterEffectDim**，且 **Animate** 属性设为 **True** 才能看到 **DimColor** 属性的效果。另外，如果指定形状是要在幻灯片上创建的唯一项目或最后一个项目，则该形状无法变暗。若要更改幻灯片上形状的创建顺序，请使用 **AnimationOrder** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前演示文稿中添加一张幻灯片。该幻灯片包含一个标题和一个三项目的列表。本示例设置标题和列表在创建后变暗，并设置它们变暗的颜色。*/ function test(){ 　　　　let myShape = Application.ActivePresentation.Slides.Add(2, ppLayoutText).Shapes 　　　　let Shape1 = myShape.Item(1)  　　　　   Shape1.TextFrame.TextRange.Text = "Sample title" 　　　　let Ani1 = Shape1.AnimationSettings  　　　　   Ani1.TextLevelEffect = ppAnimateByAllLevels   　　　　  Ani1.AfterEffect = ppAfterEffectDim   　　　　  Ani1.DimColor.SchemeColor = ppShadow  　　　　   Ani1.Animate = true 　　　　let Shape2 = myShape.Item(2)    　　　　 Shape2.TextFrame.TextRange.Text = "Item one" + "\r" + "Item two" 　　　　let Ani2 = Shape2.AnimationSettings 　　　　    Ani2.TextLevelEffect = ppAnimateByFirstLevel   　　　　  Ani2.AfterEffect = ppAfterEffectDim 　　　　    Ani2.DimColor.RGB = 100, 150, 130   　　　　  Ani2.Animate = true }` |

#### **AnimationSettings.EntryEffect**

对于 **AnimationSettings** 对象，此属性返回或设置应用于指定形状中动画的特殊效果。对于 **SlideShowTransition** 对象，此属性返回或设置应用于指定幻灯片切换的特殊效果。**PpEntryEffect** 类型，可读写。

**语法**

**express.EntryEffect**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

如果将指定形状的 **TextLevelEffect** 属性设为 **ppAnimateLevelNone**（默认值）或将 **Animate** 属性设为 **False**，您将看不到通过 **EntryEffect** 属性应用的特殊效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前演示文稿中添加一张标题幻灯片，并设置该标题在幻灯片放映中动画时从右侧飞入。*/ function test(){ 　　　　let myShape = Application.ActivePresentation.Slides.Add(1, ppLayoutTitleOnly).Shapes.Item(1)    　　　　 myShape.TextFrame.TextRange.Text = "Sample title" 　　　　let Ani = myShape.AnimationSettings 　　　　    Ani.TextLevelEffect = ppAnimateByAllLevels  　　　　   Ani.EntryEffect = ppEffectFlyFromRight     　　　　Ani.Animate = true }` |

#### **AnimationSettings.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **AnimationSettings** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **AnimationSettings.PlaySettings**

返回一个 **PlaySettings** 对象，该对象包含指定的媒体剪辑在幻灯片放映中的播放方式的信息。

**语法**

**express.PlaySettings**

*express*   一个代表 **AnimationSettings** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在当前演示文稿的第一张幻灯片中插入名为 Clock.avi 的影片，设置它在幻灯片切换后自动播放，并指定该影片对象在不播放时隐藏。*/ function test(){ 　　　　let ole = Application.ActivePresentation.Slides.Item(1).Shapes.AddOLEObject(10, 10, 250, 250, undefined, "c:\\winnt\\Clock.avi") 　　　　let ps = ole.AnimationSettings.PlaySettings 　　　　ps.PlayOnEntry = true 　　　　ps.HideWhileNotPlaying = true }` |

#### **AnimationSettings.SoundEffect**

返回一个 **SoundEffect** 对象，该对象代表切换到指定幻灯片时播放的声音。

**语法**

**express.SoundEffect**

*express*   一个代表 **AnimationSettings** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置在当前演示文稿第一张幻灯片的第一个形状演示动画时播放文件 Bass.wav。*/ function test(){ 　　　　let Ani = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).AnimationSettings 　　　　Ani.Animate = true 　　　　Ani.TextLevelEffect = ppAnimateByAllLevels 　　　　Ani.SoundEffect.ImportFromFile("c:\\bass.wav") }` |

#### **AnimationSettings.TextLevelEffect**

设置或返回一个 **PpTextLevelEffect** 常量，该常量表示指定形状中的文本是否由第一级段落、第二级段落或其他级段落（最高为第五级段落）激活动画显示。可读/写。

**语法**

**express.TextLevelEffect**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

要使 **TextLevelEffect** 属性生效，必须将 **Animate** 属性设置为 **True**。

| PpTextLevelEffect 可以是下列 PpTextLevelEffect 常量之一。 |
| --------------------------------------------------------- |
| ppAnimateByAllLevels                                      |
| ppAnimateByFifthLevel                                     |
| ppAnimateByFirstLevel                                     |
| ppAnimateByFourthLevel                                    |
| ppAnimateBySecondLevel                                    |
| ppAnimateByThirdLevel                                     |
| ppAnimateLevelMixed                                       |
| ppAnimateLevelNone                                        |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前演示文稿中添加标题幻灯片及标题文本，并将该标题的字母设置为逐个显示。*/ function test(){ 　　　　let myShape = Application.ActivePresentation.Slides.Add(1, ppLayoutTitleOnly).Shapes.Item(1)    　　　　 myShape.TextFrame.TextRange.Text = "Sample title" 　　　　let Ani = myShape.AnimationSettings   　　　　  Ani.Animate = true    　　　　 Ani.TextLevelEffect = ppAnimateByFirstLevel   　　　　  Ani.TextUnitEffect = ppAnimateByCharacter    　　　　 Ani.EntryEffect = ppEffectFlyFromLeft }` |

#### **AnimationSettings.TextUnitEffect**

返回或设置一个 **PpTextUnitEffect** 常量，该常量指示指定形状中的文本是按段落、按字或是按字母被赋予动画效果。可读/写。

**语法**

**express.TextUnitEffect**

*express*   一个代表 **AnimationSettings** 对象的变量。

**说明**

| PpTextUnitEffect 可以是下列 PpTextUnitEffect 常量之一。 |
| ------------------------------------------------------- |
| ppAnimateByCharacter                                    |
| ppAnimateByParagraph                                    |
| ppAnimateByWord                                         |
| ppAnimateUnitMixed                                      |

为使 **TextUnitEffect** 属性设置生效，必须将指定形状的 **TextLevelEffect** 属性设置为除 **ppAnimateLevelNone** 或 **ppAnimateByAllLevels** 之外的值，并且必须将 **Animate** 属性设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前演示文稿中添加标题幻灯片及标题文本，并将该标题的字母设置为逐个显示。*/ function test(){ 　　　　let myShape = Application.ActivePresentation.Slides.Add(1, ppLayoutTitleOnly).Shapes.Item(1)   　　　　  myShape.TextFrame.TextRange.Text = "Sample title" 　　　　let Ani = myShape.AnimationSettings    　　　　 Ani.Animate = true   　　　　  Ani.TextLevelEffect = ppAnimateByFirstLevel    　　　　 Ani.TextUnitEffect = ppAnimateByCharacter    　　　　 Ani.EntryEffect = ppEffectFlyFromLeft }` |

适用环境：web

适用平台：windows/linux