**FilterEffect**



代表动画动作的筛选效果。

**说明**

使用 **AnimationBehavior** 对象的 **FilterEffect** 属性返回 **FilterEffect** 对象。使用 **FilterEffect** 对象的 **Reveal**、**SubType** 和 **Type** 属性可以更改筛选效果。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reveal**      | 设置或返回一个 **MsoTriState** 常量，该常量决定嵌入对象的显示方式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Subtype**     | 设置或返回一个 **MsoAnimFilterEffectSubtype** 常量，该常量用于决定筛选效果的子类型。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**        | 返回或设置一个 **MsoAnimType** 常量，该常量代表动画的类型。可读/写。 |

**成员属性**

#### **FilterEffect.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **FilterEffect** 对象的变量。

**说明**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function AddAndSave(pptPre) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **FilterEffect.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **FilterEffect** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **FilterEffect.Reveal**

设置或返回一个 **MsoTriState** 常量，该常量决定嵌入对象的显示方式。可读/写。

**语法**

**express.Reveal**

*express*   一个代表 **FilterEffect** 对象的变量。

**说明**

当筛选效果类型为 **msoAnimFilterEffectTypeWipe** 时，如果将 **Reveal** 属性的值设置为 **msoTrue**，则会显示形状。设置为 **msoFalse** 值将使对象消失。换言之，如果将筛选设为擦除，当 Reveal 属性为真，将得到向内擦除的效果，当 Reveal 属性为假，就会得到向外擦除的效果。

| MsoTriState 可以是下列 MsoTriState 常量之一。      |
| -------------------------------------------------- |
| msoCTrue 不适用于此属性。                          |
| msoFalse 批注、修订和个人信息保留在演示文稿中。    |
| msoTriStateMixed 不适用于此属性。                  |
| msoTriStateToggle 不适用于此属性。                 |
| msoTrue 在保存演示文稿时删除批注、修订和个人信息。 |

**示例**

以下示例在当前演示文稿的第一张幻灯片中添加一个形状并设置一个筛选效果动画动作。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function ChangeFilterEffect() {     let sldFirst = Application.ActivePresentation.Slides.Item(1)     let shpHeart = sldFirst.Shapes.AddShape(msoShapeHeart,100,100,100,100)     let effNew = sldFirst.TimeLine.MainSequence.AddEffect(shpHeart,msoAnimEffectChangeFillColor,undefined,msoAnimTriggerAfterPrevious)     let bhvEffect = effNew.Behaviors.Add(msoAnimTypeFilter) 										      bhvEffect.FilterEffect.Type = msoAnimFilterEffectTypeWipe     bhvEffect.FilterEffect.Subtype = msoAnimFilterEffectSubtypeUp     bhvEffect.FilterEffect.Reveal = msoTrue }` |

#### **FilterEffect.Subtype**

设置或返回一个 **MsoAnimFilterEffectSubtype** 常量，该常量用于决定筛选效果的子类型。可读/写。

**语法**

**express.Subtype**

*express*   一个代表 **FilterEffect** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在当前演示文稿的第一张幻灯片中添加一个形状并设置一个筛选效果动画动作。*/ function ChangeFilterEffect() {     let sldFirst = Application.ActivePresentation.Slides.Item(1)     let shpHeart = sldFirst.Shapes.AddShape(msoShapeHeart,100,100,100,100)     let effNew = sldFirst.TimeLine.MainSequence.AddEffect(shpHeart,msoAnimEffectChangeFillColor,undefined,msoAnimTriggerAfterPrevious)     let bhvEffect = effNew.Behaviors.Add(msoAnimTypeFilter) 										      bhvEffect.FilterEffect.Type = msoAnimFilterEffectTypeWipe     bhvEffect.FilterEffect.Subtype = msoAnimFilterEffectSubtypeUp     bhvEffect.FilterEffect.Reveal = msoTrue }` |

#### **FilterEffect.Type**

返回或设置一个 **MsoAnimType** 常量，该常量代表动画的类型。可读/写。

**语法**

**express.Type**

*express*   一个代表 **FilterEffect** 对象的变量。

**说明**

| MsoAnimType 可以是下列 MsoAnimType 常量之一。 |
| --------------------------------------------- |
| MsoAnimTypeColor                              |
| MsoAnimTypeMixed                              |
| MsoAnimTypeMotion                             |
| MsoAnimTypeNone                               |
| MsoAnimTypeProperty                           |
| MsoAnimTypeRoatation                          |
| MsoAnimTypeScale                              |
| MsoAnimTypeTransition                         |

适用环境：web

适用平台：windows/linux