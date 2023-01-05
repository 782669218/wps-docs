**Design**



代表单张幻灯片设计模板。**Design** 对象是 **Designs** 和 **SlideRange** 集合以及 **Master** 和 **Slide** 对象的成员。

**说明**

使用 **Master**、**Slide** 或 **SlideRange** 对象的 **Design** 属性访问 **Design** 对象，例如：

| 示例代码                                            |
| --------------------------------------------------- |
| `Application.ActivePresentation.SlideMaster.Design` |

| 示例代码                                               |
| ------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Design` |

| 示例代码                                   |
| ------------------------------------------ |
| `ActivePresentation.Slides.Range().Design` |

分别使用 **Designs** 集合的 **Add**、**Item**、**Clone** 或 **Load** 方法添加、引用、克隆或加载一个 **Design** 对象。例如，若要添加一个设计模板，请使用 `ActivePresentation.Designs.Add designName:="MyDesign"`

**方法**

|                                                              | 名称       | 说明                                                         |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveTo** | 将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回 **Long** 类型值，该值代表动画效果或设计的索引号。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回或设置指定对象的名称。**String** 类型，可读/写。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Preserved**   | 设置或返回一个 **MsoTriState** 类型的常量，该常量代表是否保护设计母版以避免被更改。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideMaster** | 返回一个 **Master** 对象，该对象代表幻灯片母版。             |

**成员方法**

#### **Design.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **Design** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

#### **Design.MoveTo**

将指定对象移动到同一集合中的指定位置，并适当地对集合中所有其他项目进行重新编号。

**语法**

**express.MoveTo(toPos)**

*express*   一个代表 **Design** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *toPos*  | 必选          | **Long**     | 动画效果要移动到的索引。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在指定形状的动画效果集合中将一个动画效果移动到第二个动画效果的位置上。*/ function MoveEffect() {     let sldFirst = Application.ActivePresentation.Slides.Item(1)     let shpFirst = sldFirst.Shapes.Item(1)     let effAdd = sldFirst.TimeLine.MainSequence.AddEffect(shpFirst, msoAnimEffectBlinds)     effAdd.MoveTo(2) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿中的第二张幻灯片移动到第一张幻灯片的位置上。*/  Application.ActivePresentation.Slides.Item(2).MoveTo(1)` |

**成员属性**

#### **Design.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Design** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **Design.Index**

返回 **Long** 类型值，该值代表动画效果或设计的索引号。只读。

**语法**

**express.Index**

*express*   一个代表 **Design** 对象的变量。

**示例**

以下示例显示第一张幻灯片的主动画序列中所有效果的名称和索引号。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ` alert("Effect Name: " + Application.ActivePresentation.Slides.Item(1).DisplayName + "\n" + "Effect Index: " + Application.ActivePresentation.Slides.Item(1).Index)` |

#### **Design.Name**

返回或设置指定对象的名称。**String** 类型，可读/写。

**语法**

**express.Name**

*express*   一个代表 **Design** 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。例如，如果某个形状的 **Name** 属性值为“Rectangle 2”，则 `.Shapes("Rectangle 2")` 将返回对该形状的引用。

#### **Design.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Design** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **Design.Preserved**

设置或返回一个 **MsoTriState** 类型的常量，该常量代表是否保护设计母版以避免被更改。可读/写。

**语法**

**express.Preserved**

*express*   一个代表 **Design** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue 不适用于此属性。                     |
| msoFalse 设计母版尚未保护并且可进行编辑。     |
| msoTriStateMixed 不适用于此属性。             |
| msoTriStateToggle 不适用于此属性。            |
| msoTrue 设计母版已保护且无法进行编辑。        |

**示例**

下面的代码行锁定并保护第一个设计母版。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ` Application.ActivePresentation.Designs.Item(1).Preserved = msoTrue` |

#### **Design.SlideMaster**

返回一个 **Master** 对象，该对象代表幻灯片母版。

**语法**

**express.SlideMaster**

*express*   一个代表 **Design** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置当前演示文稿的幻灯片母版的背景图案。*/ function test(){ 　　　　Application.ActivePresentation.SlideMaster.Background.Fill.PresetTextured(msoTextureGreenMarble) }` |

适用环境：web

适用平台：windows/linux