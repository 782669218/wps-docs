**ShapeRange**



代表一个形状范围，即某个文档中的一组形状。形状范围可以只包含文档中的一个形状，也可以包含该文档中的所有形状。

**说明**

您可以使用任何所需的形状（从文档的所有形状中进行选择，或从选定内容的所有形状中进行选择）构造形状范围。例如，可以构造一个 **ShapeRange** 集合，其中包含文档中的前三个形状、文档中所有选定的形状，或文档中所有的任意多边形。

有关如何使用一个形状或同时使用多个形状的概述，请参阅 使用形状（图形对象）。

以下示例说明如何执行下列操作：

- 返回一组通过名称或索引序号指定的形状。

- 返回文档中全部或部分选定的形状。

**方法**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Align**             | 对齐指定形状区域中的形状。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Apply**             | 适用于指定的形状格式，该格式已使用 [PickUp](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#Shape.PickUp) 方法复制。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**              | 将指定对象复制到剪贴板。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**               | 删除指定对象并将其放到剪贴板中。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**            | 删除指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Duplicate**         | 创建指定的 ShapeRange 对象的副本，将新的形状或形状范围添加到 Shapes 集合中，然后返回新的 ShapeRange 对象。副本对象位于 Shapes 集合末尾。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Flip**              | 绕水平或垂直轴翻转指定形状。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Group**             | 将指定区域中的形状形成一组。以单个 Shape 对象返回分组后的形状。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementLeft**     | 以指定点数水平移动指定形状。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementRotation** | 按指定度数改变指定形状绕 z 轴的旋转量。使用 Rotation 属性设置形状的绝对旋转量。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementTop**      | 以指定点数垂直移动指定形状。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**              | 从指定集合中返回单个对象。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PickUp**            | 复制指定形状的格式。用 **Apply** 方法可将复制的格式应用于其他形状。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScaleHeight**       | 以指定的比例缩放图形高度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于除图片和 OLE 对象以外的其他图形来说，缩放总是相对于当前高度而言。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScaleWidth**        | 根据指定的系数缩放图形宽度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于图片和 OLE 对象以外的其他图形来说，总是相对于当前宽度进行缩放。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**            | 选择指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Ungroup**           | 取消指定形状或者形状区域中任意组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。取消组合后的形状以单个 ShapeRange 对象的形式返回。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ZOrder**            | 将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-顺序中的位置）。 |

**属性**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActionSettings**    | 返回一个 ActionSettings 对象，该对象包含在幻灯片放映期间，当用户在指定形状或文本区域内单击或移动鼠标时所产生的动作的信息。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**       | 返回一个 Application 对象，该对象表示指定对象的创建者。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Chart**             | 返回当前 Shape 对象的 Chart 对象。只读。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**             | 返回指定集合中的对象数目。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fill**              | 返回一个 FillFormat 对象，该对象包含指定形状的填充格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Glow**              | 返回指定形状范围的发光格式。只读 msoGlowType 类型。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GroupItems**        | 返回一个 GroupShapes 对象，该对象代表指定形状组中的单个形状。使用 GroupShapes 对象的 Item 方法可返回形状组中的单个形状。适用于代表组合形状的 Shape 或 ShapeRange 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasChart**          | 返回指定的对象代表的形状区域是否包含图表。只读。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasTable**          | 返回指定的形状是否为表。只读。MsoTriState 类型。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasTextFrame**      | 返回指定形状是否有文本框。只读。MsoTriState 类型。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**            | 以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**                | 返回一个 Long 类型值，该值标识形状或形状范围。只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**              | 返回或设置一个 Single 类型值，该值代表从形状范围中最左侧形状的左边缘到幻灯片左边缘的距离（以磅为单位）。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Line**              | 返回一个 LineFormat 对象，该对象包含指定形状的线条格式属性。（对于线条来说，LineFormat 对象代表线条本身；而对于带有边框的形状来说，LineFormat 对象代表边框。）只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MediaType**         | 返回 OLE 媒体类型。只读。PpMediaType 类型。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**              | 创建形状时，WPP 自动以 ShapeType Number 的形式为其分配一个名称，其中 ShapeType 指明形状或自选图形的类型，Number 是在幻灯片上形状的集合中具有唯一性的整数。例如，为幻灯片上的形状自动生成的名称可以为 Placeholder 1、Oval 2 和 Rectangle 3。为了避免与自动分配的名称发生冲突，对用户定义的名称请不要使用 ShapeType Number 形式，其中 ShapeType 为一个用于自动生成名称的值，Number 为任意的正整数。形状范围必须至少包含一个形状。String 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureFormat**     | 返回一个 PictureFormat 对象，该对象包含指定形状的图片格式属性。适用于代表图片或 OLE 对象的 Shape 或 ShapeRange 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PlaceholderFormat** | 返回一个 PlaceholderFormat 对象，该对象包含占位符特有的属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**            | 返回一个只读的 ShadowFormat 对象，该对象包含指定形状的阴影格式属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShapeStyle**        | 设置或返回指定对象的形状样式索引                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Table**             | 返回一个 Table 对象，该对象代表某个形状或形状区域内的一个表格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextEffect**        | 返回一个 TextEffectFormat 对象，该对象包含指定形状的文本效果格式属性。适用于代表艺术字的 Shape 或 ShapeRange 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame**         | 返回一个 TextFrame 对象，该对象包含指定形状或母版文本样式的对齐方式和定位属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame2**        | 返回与指定的 ShapeRange 对象相关联的 TextFrame2 对象，前者包含指定的形状范围的对齐和定位属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThreeD**            | 返回一个 ThreeDFormat 对象，该对象包含指定形状的三维效果格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**               | 返回或设置一个 Single 类型值，该值代表形状范围内最上方形状的上边缘到文档上边缘的距离。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**              | 返回一个 **MsoShapeType** 常量，该常量代表单个形状或形状范围内多个形状的类型。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**           | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**             | 以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ZOrderPosition**    | 返回指定的形状在 z-顺序中的位置。只读。Long 类型。           |

**成员方法**

#### **ShapeRange.Align**

对齐指定形状区域中的形状。

**语法**

**express.Align(AlignCmd, RelativeTo)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**    | **说明**                               |
| ------------ | ------------- | --------------- | -------------------------------------- |
| *AlignCmd*   | 必选          | **MsoAlignCmd** | 指定某个指定形状范围中形状的对齐方式。 |
| *RelativeTo* | 必选          | **MsoTriState** | 决定形状是否与幻灯片边缘对齐。         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中指定范围的所有形状的左边与该范围内最左边的形状的左边对齐*/ function  test(){     let myDocument = Application.ActivePresentation.Slides.Item(3)     myDocument.Shapes.Range().Align(msoAlignLefts, msoFalse) }` |

#### **ShapeRange.Apply**

适用于指定的形状格式，该格式已使用 [PickUp](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#Shape.PickUp) 方法复制。

**语法**

**express.Apply()**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例复制 myDocument 上第一个形状的格式，然后将该格式应用于第二个形状*/ function test(){ 	let myDocument = Application.ActivePresentation.Slides.Item(1) 	myDocument.Shapes.Item(1).PickUp() 	myDocument.Shapes.Item(2).Apply() }` |

#### **ShapeRange.Copy**

将指定对象复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿的第一张幻灯片复制到剪贴板。*/ Application.ActivePresentation.Slides.Item(1).Copy()  /*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy()` |

#### **ShapeRange.Cut**

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut()**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中*/  Application.Windows.Item(1).Selection.Cut()  /*本示例从当前演示文稿的第一张幻灯片中删除第一个和第二个形状，并将它们的副本放到剪贴板，然后将该副本粘贴到第二张幻灯片*/ function test(){   ActivePresentation.Slides.Item(1).Shapes.Range([1, 2]).Cut()   ActivePresentation.Slides.Item(2).Shapes.Paste() }` |

#### **ShapeRange.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除第一个窗口中选中的对象 Application.Windows.Item(1).Selection.Delete()` |

#### **ShapeRange.Duplicate**

创建指定的 ShapeRange 对象的副本，将新的形状或形状范围添加到 Shapes 集合中，然后返回新的 ShapeRange 对象。副本对象位于 Shapes 集合末尾。

**语法**

**express.Duplicate()**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Flip**

绕水平或垂直轴翻转指定形状。

**语法**

**express.Flip(FlipCmd)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**   | **说明**                       |
| --------- | ------------- | -------------- | ------------------------------ |
| *FlipCmd* | 必选          | **MsoFlipCmd** | 指定形状是水平翻转还是垂直翻转 |

**说明**

MsoFlipCmd 可以是下列 MsoFlipCmd 常量之一。
msoFlipHorizontal
msoFlipVertical

#### **ShapeRange.Group**

将指定区域中的形状形成一组。以单个 Shape 对象返回分组后的形状。

**语法**

**express.Group()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

由于一组形状作为单个形状处理，所以创建和分解形状组将更改 Shapes 集合中的项目数，并更改集合中受影响的项目之后的项目索引号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将两个形状添加到第一张幻灯片，组合两个新形状，为该组设置填充，旋转该组并将其发送到绘图层后*/ function test(){ 	let myShapes =  Application.ActivePresentation.Slides.Item(1).Shapes  	myShapes.AddShape(msoShapeCan, 50, 10, 100, 200).Name = "shpOne" 	myShapes.AddShape(msoShapeCube, 150, 250, 100, 200).Name = "shpTwo"  	let grouped = myShapes.Range(["shpOne", "shpTwo"]).Group() 	grouped.Fill.PresetTextured(msoTextureBlueTissuePaper) 	grouped.Rotation = 45 	grouped.ZOrder(msoSendToBack) }` |

#### **ShapeRange.IncrementLeft**

以指定点数水平移动指定形状。

**语法**

**express.IncrementLeft(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 以点为单位指定形状水平移动的距离。正值表示将形状向右移动，负值表示将形状向左移动。 |

#### **ShapeRange.IncrementRotation**

按指定度数改变指定形状绕 z 轴的旋转量。使用 Rotation 属性设置形状的绝对旋转量。

**语法**

**express.IncrementRotation(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 以度为单位指定形状的水平旋转量。正值表示使形状按顺时针方向旋转；负值表示使形状按逆时针方向旋转。 |

**说明**

要将三维形状绕 X 轴或 Y 轴旋转，请分别使用 IncrementRotationX 方法或 IncrementRotationY 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例复制myDocument上的第一个形状，并设置复制的填充，再向右移70磅、向上移50磅，然后顺时针旋转30度角*/ function test(){ 	let myShape = ActivePresentation.Slides.Item(1).Shapes.Item(1).Duplicate() 	myShape.Fill.PresetTextured(msoTextureGranite) 	myShape.IncrementLeft(70) 	myShape.IncrementTop(-50) 	myShape.IncrementRotation(30) }` |

#### **ShapeRange.IncrementTop**

以指定点数垂直移动指定形状。

**语法**

**express.IncrementTop(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 以点为单位指定形状对象的垂直移动距离。正值表示将形状向下移动，负值表示将形状向上移动。 |

**说明**

要将三维形状绕 X 轴或 Y 轴旋转，请分别使用 IncrementRotationX 方法或 IncrementRotationY 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例复制 myDocument 上的第一个形状，并设置复制的填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度角*/ function test(){ 	let myShape = ActivePresentation.Slides.Item(1).Shapes.Item(1).Duplicate() 	myShape.Fill.PresetTextured(msoTextureGranite) 	myShape.IncrementLeft(70) 	myShape.IncrementTop(-50) 	myShape.IncrementRotation(30) }` |

#### **ShapeRange.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Variant**  | 集合中要返回的单个对象的名称或索引号。 |

#### **ShapeRange.PickUp**

复制指定形状的格式。用 **Apply** 方法可将复制的格式应用于其他形状。

**语法**

**express.PickUp()**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例复制 myDocument 上第一个形状的格式，然后将该格式应用于第二个形状*/ function test(){ 	let myDocument =Application.ActivePresentation.Slides.Item(1)  	myDocument.Shapes.Item(1).PickUp() 	myDocument.Shapes.Item(2).Apply() }` |

#### **ShapeRange.ScaleHeight**

以指定的比例缩放图形高度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于除图片和 OLE 对象以外的其他图形来说，缩放总是相对于当前高度而言。

**语法**

**express.ScaleHeight(Factor, RelativeToOriginalSize, fScale)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**     | **说明**                                                     |
| ------------------------ | ------------- | ---------------- | ------------------------------------------------------------ |
| *Factor*                 | 必选          | **Single**       | 指定形状调整后的高度与当前或原始高度的比例。例如，若要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选          | **MsoTriState**  | 指定是否相对于形状的当前或原始尺寸来缩放形状。               |
| *fScale*                 | 可选          | **MsoScaleFrom** | 在缩放形状时，形状中位置不变的部分。                         |

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。
msoCTrue
msoFalse：相对于形状的当前尺寸缩放。
msoTriStateMixed
msoTriStateToggle
msoTrue 相对于原始大小缩放该形状。仅当指定形状是图片或 OLE 对象时，才为此参数指定 msoTrue。
MsoScaleFrom 可以是下列 MsoScaleFrom 常量之一。
msoScaleFromBottomRight
msoScaleFromMiddle
msoScaleFromTopLeft：默认值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%*/ function test(){ 	let myDocument = Application.ActivePresentation.Slides.Item(1)  	for(let i = 1; i <= myDocument.Shapes.Count; i++){ 		switch(myDocument.Shapes.Item(i).Type){ 			case msoEmbeddedOLEObject: 			case msoLinkedOLEObject: 			case msoOLEControlObject: 			case msoLinkedPicture: 			case msoPicture: 				myDocument.Shapes.Item(i).ScaleHeight(1.75, msoTrue) 				myDocument.Shapes.Item(i).ScaleWidth(1.75, msoTrue) 				break 			default: 				myDocument.Shapes.Item(i).ScaleHeight(1.75, msoFalse) 				myDocument.Shapes.Item(i).ScaleWidth(1.75, msoFalse) 				break 		} 	} }` |

#### **ShapeRange.ScaleWidth**

根据指定的系数缩放图形宽度。对于图片和 OLE 对象来说，可指明图形缩放是根据原尺寸还是当前尺寸。对于图片和 OLE 对象以外的其他图形来说，总是相对于当前宽度进行缩放。

**语法**

**express.ScaleWidth(Factor, RelativeToOriginalSize, fScale)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**     | **说明**                                                     |
| ------------------------ | ------------- | ---------------- | ------------------------------------------------------------ |
| *Factor*                 | 必选          | **Single**       | 指定形状调整后的宽度与当前或原始宽度的比例。例如，若要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选          | **MsoTriState**  | 指定是否相对于形状的当前或原始尺寸来缩放形状。               |
| *fScale*                 | 可选          | **MsoScaleFrom** | 在缩放形状时，形状中位置不变的部分。                         |

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。
msoCTrue
msoFalse：相对于形状的当前尺寸缩放。
msoTriStateMixed
msoTriStateToggle
msoTrue 相对于原始大小缩放该形状。仅当指定形状是图片或 OLE 对象时，才为此参数指定 msoTrue。
MsoScaleFrom 可以是下列 MsoScaleFrom 常量之一。
msoScaleFromBottomRight
msoScaleFromMiddle
msoScaleFromTopLeft：默认值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 上的所有图片和 OLE 对象放大至原高度和宽度的 175%，将所有其他形状放大至当前高度和宽度的 175%*/ function test(){ 	let myDocument = Application.ActivePresentation.Slides.Item(1)  	for(let i = 1; i <= myDocument.Shapes.Count; i++){ 		switch(myDocument.Shapes.Item(i).Type){ 			case msoEmbeddedOLEObject: 			case msoLinkedOLEObject: 			case msoOLEControlObject: 			case msoLinkedPicture: 			case msoPicture: 				myDocument.Shapes.Item(i).ScaleHeight(1.75, msoTrue) 				myDocument.Shapes.Item(i).ScaleWidth(1.75, msoTrue) 				break 			default: 				myDocument.Shapes.Item(i).ScaleHeight(1.75, msoFalse) 				myDocument.Shapes.Item(i).ScaleWidth(1.75, msoFalse) 				break 		} 	} }` |

#### **ShapeRange.Select**

选择指定的对象。

**语法**

**express.Select(Replace)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**    | **说明**                                   |
| --------- | ------------- | --------------- | ------------------------------------------ |
| *Replace* | 可选          | **MsoTriState** | 指定该选择是否替换之前所做的任何一项选择。 |

**说明**

如果做出不适用于当前视图的选择，此代码将失败。例如，可以在幻灯片浏览视图中而不能在幻灯片视图中选择幻灯片。
MsoTriState 可以是下列 MsoTriState 常量之一。
msoCTrue
msoFalse 将该选择添加到前一项选择中。
msoTriStateMixed
msoTriStateToggle
msoTrue 默认值。该选择替换之前所做的任何一项选择。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择活动演示文稿第一张幻灯片上的第一个和第三个形状*/ Application.ActivePresentation.Slides.Item(1).Shapes.Range([1, 3]).Select()` |

#### **ShapeRange.Ungroup**

取消指定形状或者形状区域中任意组合形状的组合。取消指定形状或形状区域中图片和 OLE 对象的组合。取消组合后的形状以单个 ShapeRange 对象的形式返回。

**语法**

**express.Ungroup()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

由于将一组形状作为单个形状来处理，对形状进行组合和取消组合将会改变 Shapes 集合中项目的数量，并且改变集合中受影响的项目之后的项目索引编号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例取消 myDocument 中的所有形状组合，并分解所有的图片或 OLE 对象*/ function test(){ 	let myDocument = Application.ActivePresentation.Slides.Item(1)  	for(let i = 1; i <= myDocument.Shapes.Count; i++){ 		myDocument.Shapes.Item(i).Ungroup() 	} }  /*本示例取消 myDocument 中的所有形状组合，但不分解幻灯片中的图片或 OLE 对象*/ function test(){ 	let myDocument = Application.ActivePresentation.Slides.Item(1)  	for(let i = 1; i <= myDocument.Shapes.Count; i++){ 		if(myDocument.Shapes.Item(i).Type == msoGroup) { 		   myDocument.Shapes.Item(i).Ungroup() 		} 	} }` |

#### **ShapeRange.ZOrder**

将指定的形状移到集合中其他形状的前面或后面（即更改该形状在 z-顺序中的位置）。

**语法**

**express.ZOrder(ZorderCmd)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**     | **说明**                                 |
| ----------- | ------------- | ---------------- | ---------------------------------------- |
| *ZorderCmd* | 必选          | **MsoZOrderCmd** | 指定针对其他形状将指定形状移动到的位置。 |

**说明**

MsoZOrderCmd 可以是下列 MsoZOrderCmd 常量之一。
msoBringForward
msoBringInFrontOfText 仅用于 WPS。
msoBringToFront
msoSendBackward
msoSendBehindText 仅用于 WPS。
msoSendToBack
可用 ZOrderPosition 属性确定形状在 z-顺序中的当前位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加椭圆，如果幻灯片中有至少一个其他形状，则将该椭圆放置在 z-顺序中倒数第二的位置上*/ function test(){ 	let myDocument = Application.ActivePresentation.Slides.Item(1) 	let myShape = myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300)  	while(myShape.ZOrderPosition > 2){ 		myShape.ZOrder(msoSendBackward) 	} }` |

**成员属性**

#### **ShapeRange.ActionSettings**

返回一个 ActionSettings 对象，该对象包含在幻灯片放映期间，当用户在指定形状或文本区域内单击或移动鼠标时所产生的动作的信息。只读。

**语法**

**express.ActionSettings**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置在当前演示文稿中第二张幻灯片的第一个形状中单击或移动鼠标时产生的动作*/ function test(){     let myShape = Application.ActivePresentation.Slides.Item(2).Shapes.Item(1)     myShape.ActionSettings(ppMouseClick).Action = ppActionLastSlide     myShape.ActionSettings(ppMouseOver).SoundEffect.Name = "applause.wav" }` |

#### **ShapeRange.Application**

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/ function test() {     let pptPres = Application.ActivePresentation     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称*/ function test(){     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes      for(let i = 1; i <= myShapes.Count; i++){         if(myShapes.Item(i).Type == msoLinkedOLEObject){             MsgBox(myShapes.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **ShapeRange.Chart**

返回当前 Shape 对象的 Chart 对象。只读。

**语法**

**express.Chart**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口*/ function test(){     for(let i = 2;i <= Application.Windows.Count;i++){         Application.Windows.Item(2).Close()     } }` |

#### **ShapeRange.Fill**

返回一个 FillFormat 对象，该对象包含指定形状的填充格式属性。只读。

**语法**

**express.Fill**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变*/ function test(){     let myDocument = Application.ActivePresentation.Slides.Item(1)     let filler = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill      filler.ForeColor.RGB = (128, 0, 0)     filler.BackColor.RGB = (170, 170, 170)     filler.TwoColorGradient(msoGradientHorizontal, 1) }` |

#### **ShapeRange.Glow**

返回指定形状范围的发光格式。只读 msoGlowType 类型。

**语法**

**express.Glow**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.GroupItems**

返回一个 GroupShapes 对象，该对象代表指定形状组中的单个形状。使用 GroupShapes 对象的 Item 方法可返回形状组中的单个形状。适用于代表组合形状的 Shape 或 ShapeRange 对象。只读。

**语法**

**express.GroupItems**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向myDocument中添加三个三角形，将它们组合，再设置整个组的颜色，然后只更改第二个三角形的颜色*/ function test(){     let myDocument = Application.ActivePresentation.Slides.Item(1)      myDocument.Shapes.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"     myDocument.Shapes.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"     myDocument.Shapes.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"      let grouped = myDocument.Shapes.Range(["shpOne", "shpTwo", "shpThree"]).Group()     grouped.Fill.PresetTextured(msoTextureBlueTissuePaper)     grouped.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble) }` |

#### **ShapeRange.HasChart**

返回指定的对象代表的形状区域是否包含图表。只读。

**语法**

**express.HasChart**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.HasTable**

返回指定的形状是否为表。只读。MsoTriState 类型。

**语法**

**express.HasTable**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。
msoCTrue
msoFalse
msoTriStateMixed
msoTriStateToggle
msoTrue：指定的形状是表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例检查当前选定的形状是否为表格。如果是表格，此代码将第一列的宽度设置为 1 英寸*/ function test(){     let myRange = ActiveWindow.Selection.ShapeRange     if(myRange.HasTable == msoTrue){         myRange.Table.Columns.Item(1).Width = 72     } }` |

#### **ShapeRange.HasTextFrame**

返回指定形状是否有文本框。只读。MsoTriState 类型。

**语法**

**express.HasTextFrame**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。
msoCTrue
msoFalse
msoTriStateMixed
msoTriStateToggle
msoTrue：指定形状有文本框，因此可包含文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的例子从第一张幻灯片上所有包含文本框的形状中提取文本，然后将这些形状的名称及其所包含的文本保存在一个数组中*/ function test(){     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes      if(myShapes.Count > 1){         let numTextShapes = 0         let shpTextArray = []          for(let i = 1; i <= myShapes.Count; i++){             if(myShapes.Item(i).HasTextFrame){                 numTextShapes++                 shpTextArray[numTextShapes] = []                 shpTextArray[numTextShapes][0] = myShapes.Item(i).Name                 shpTextArray[numTextShapes][1] = myShapes.Item(i).TextFrame.TextRange.Text             }         }     } }` |

#### **ShapeRange.Height**

以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Height**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

一个 HeightShape 对象的 ShapeHeight 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将第二个文档窗口的高度设置为应用程序窗口高度的一半*/ Application.Windows.Item(2).Height = Application.Height / 2  /*以下示例将指定表格中第二行的高度设为 100 磅*/ Application.ActivePresentation.Slides.Item(2).Shapes.Item(2).Table.Rows.Item(2).Height = 100 ` |

#### **ShapeRange.Id**

返回一个 Long 类型值，该值标识形状或形状范围。只读。

**语法**

**express.Id**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向当前演示文稿添加一个新形状，再根据 ID 属性的值填充该形状*/ function test(){     let myShape = Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShape5pointStar, 100, 100, 100, 100)      switch(true){         case myShape.Id <= 500:             myShape.Fill.ForeColor.RGB = (255, 0, 0)             break         case myShape.Id <= 1000:             myShape.Fill.ForeColor.RGB = (255, 255, 0)             break         case myShape.Id <= 1500:             myShape.Fill.ForeColor.RGB = (255, 0, 255)             break         case myShape.Id <= 2000:             myShape.Fill.ForeColor.RGB = (0, 255, 0)             break         case myShape.Id <= 2500:             myShape.Fill.ForeColor.RGB = (0, 255, 255)             break         default:             myShape.Fill.ForeColor.RGB = (0, 0, 255)     } }` |

#### **ShapeRange.Left**

返回或设置一个 Single 类型值，该值代表从形状范围中最左侧形状的左边缘到幻灯片左边缘的距离（以磅为单位）。可读/写。

**语法**

**express.Left**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Line**

返回一个 LineFormat 对象，该对象包含指定形状的线条格式属性。（对于线条来说，LineFormat 对象代表线条本身；而对于带有边框的形状来说，LineFormat 对象代表边框。）只读。

**语法**

**express.Line**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加一条蓝色虚线*/ function test(){     let myDocument = Application.ActivePresentation.Slides.Item(1)     let myLine = myDocument.Shapes.AddLine(10, 10, 250, 250).Line      myLine.DashStyle = msoLineDashDotDot     myLine.ForeColor.RGB = (50, 0, 128) }` |

#### **ShapeRange.MediaType**

返回 OLE 媒体类型。只读。PpMediaType 类型。

**语法**

**express.MediaType**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

PpMediaType 可以是下列 PpMediaType 常量之一。
ppMediaTypeMixed
ppMediaTypeMovie
ppMediaTypeOther
ppMediaTypeSound

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置在放映幻灯片时，当前演示文稿内第一张幻灯片上的所有原始声音对象循环播放直到被手动停止*/ function test(){ 	let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes  	for(let i = 1; i <= myShapes.Count; i++){ 		if(myShapes.Item(i).Type == msoMedia){ 			if(myShapes.Item(i).MediaType == ppMediaTypeSound){ 				myShapes.Item(i).AnimationSettings.PlaySettings.LoopUntilStopped = true 			} 		} 	} }` |

#### **ShapeRange.Name**

创建形状时，WPP 自动以 ShapeType Number 的形式为其分配一个名称，其中 ShapeType 指明形状或自选图形的类型，Number 是在幻灯片上形状的集合中具有唯一性的整数。例如，为幻灯片上的形状自动生成的名称可以为 Placeholder 1、Oval 2 和 Rectangle 3。为了避免与自动分配的名称发生冲突，对用户定义的名称请不要使用 ShapeType Number 形式，其中 ShapeType 为一个用于自动生成名称的值，Number 为任意的正整数。形状范围必须至少包含一个形状。String 类型，可读/写。

**语法**

**express.Name**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

如果包含某对象的集合的 Item 方法采用 Variant 类型的参数，则可以将该对象的名称与 Item 方法联合使用来返回对该对象的引用。例如，如果某个形状的 Name 属性值为“Rectangle 2”，则 .Shapes.Item("Rectangle 2") 将返回对该形状的引用。

#### **ShapeRange.PictureFormat**

返回一个 PictureFormat 对象，该对象包含指定形状的图片格式属性。适用于代表图片或 OLE 对象的 Shape 或 ShapeRange 对象。只读。

**语法**

**express.PictureFormat**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置了 myDocument 中第一个形状的亮度和对比度。第一个形状必须是图片或 OLE 对象*/ function test(){ 	let myDocument = ActivePresentation.Slides.Item(1) 	myDocument.Shapes.Item(1).PictureFormat.Brightness = 0.3 	myDocument.Shapes.Item(1).PictureFormat.Contrast = .75 }` |

#### **ShapeRange.PlaceholderFormat**

返回一个 PlaceholderFormat 对象，该对象包含占位符特有的属性。只读。

**语法**

**express.PlaceholderFormat**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在当前演示文稿内第一张幻灯片的第一个占位符中添加文本（如果该占位符是水平标题占位符*/ function test(){ 	let phs = Application.ActivePresentation.Slides.Item(1).Shapes.Placeholders  	if(phs.Count > 0){ 		switch(phs.Item(1).PlaceholderFormat.Type){ 			case ppPlaceholderTitle: 				phs.Item(1).TextFrame.TextRange.Text = "Title Text" 				break 			case ppPlaceholderCenterTitle: 				phs.Item(1).TextFrame.TextRange.Text = "Centered Title Text" 				break 			default: 				MsgBox("There's no horizontal title on this slide") 		} 	} }` |

#### **ShapeRange.Shadow**

返回一个只读的 ShadowFormat 对象，该对象包含指定形状的阴影格式属性。

**语法**

**express.Shadow**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前演示文稿第一张幻灯片中添加一个有阴影的矩形。蓝色浮凸效果的阴影距该矩形的右边 3 磅，下边 2 磅*/ function test(){     let myShap = Application.ActivePresentation.Slides.Item(1).Shapes     let myShad = myShap.AddShape(msoShapeRectangle, 10, 10, 150, 90).Shadow         myShad.Type = msoShadow17         myShad.ForeColor.RGB = (0, 0, 128)         myShad.OffsetX = 3         myShad.OffsetY = 2 }` |

#### **ShapeRange.ShapeStyle**

设置或返回指定对象的形状样式索引

**语法**

**express.ShapeStyle**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Table**

返回一个 Table 对象，该对象代表某个形状或形状区域内的一个表格。只读。

**语法**

**express.Table**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将第三张幻灯片第二个形状中表格第二列的宽度设置为 60 磅*/ Application.ActivePresentation.Slides.Item(3).Shapes.Item(2).Table.Rows.Item(2).Height = 60` |

#### **ShapeRange.TextEffect**

返回一个 TextEffectFormat 对象，该对象包含指定形状的文本效果格式属性。适用于代表艺术字的 Shape 或 ShapeRange 对象。

**语法**

**express.TextEffect**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 myDocument 上第三个形状的字体样式设置为加粗（如果该形状为艺术字）*/ function test(){ 	let myDocument = Application.ActivePresentation.Slides.Item(1) 	let myShape = myDocument.Shapes.Item(3) 		if(myShape.Type == msoTextEffect) { 			myShape.TextEffect.FontBold = true 		} }` |

#### **ShapeRange.TextFrame**

返回一个 TextFrame 对象，该对象包含指定形状或母版文本样式的对齐方式和定位属性。

**语法**

**express.TextFrame**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

使用 TextFrame 对象的 TextRange 属性可返回文本框中的文本。
在应用 TextFrame 属性之前，可使用 HasTextFrame 属性来确定形状中是否包含文本框。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在 myDocument 中添加一个矩形，向该矩形中添加文本，然后设置文本框的上边距*/ function test(){ 	let myDocument = Application.ActivePresentation.Slides.Item(1) 	let myTextFrame = myDocument.Shapes  			.AddShape(msoShapeRectangle, 180, 175, 350, 140).TextFrame 		myTextFrame.TextRange.Text = "Here is some test text" 		myTextFrame.MarginTop = 10 }` |

#### **ShapeRange.TextFrame2**

返回与指定的 ShapeRange 对象相关联的 TextFrame2 对象，前者包含指定的形状范围的对齐和定位属性。只读。

**语法**

**express.TextFrame2**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.ThreeD**

返回一个 ThreeDFormat 对象，该对象包含指定形状的三维效果格式属性。只读。

**语法**

**express.ThreeD**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为应用于 myDocument 中第一个形状的三维效果设置深度、延伸色、延伸方向以及照明方向*/ function test(){ 	let myDocument =Application.ActivePresentation.Slides.Item(1) 	let myThreeD = myDocument.Shapes.Item(1).ThreeD 		myThreeD.Visible = true 		myThreeD.Depth = 50 		//RGB value for purple 		myThreeD.ExtrusionColor.RGB = (255, 100, 255) 		myThreeD.SetExtrusionDirection(msoExtrusionTop) 		myThreeD.PresetLightingDirection = msoLightingLeft }` |

#### **ShapeRange.Top**

返回或设置一个 Single 类型值，该值代表形状范围内最上方形状的上边缘到文档上边缘的距离。可读/写。

**语法**

**express.Top**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Type**

返回一个 **MsoShapeType** 常量，该常量代表单个形状或形状范围内多个形状的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

MsoShapeType 可以是下列 MsoShapeType 常量之一。
msoAutoShape
msoCallout
msoCanvas
msoChart
msoComment
msoDiagram
msoEmbeddedOLEObject
msoFormControl
msoFreeform
msoGroup
msoLine
msoLinkedOLEObject
msoLinkedPicture
msoMedia
msoOLEControlObject
msoPicture
msoPlaceholder
msoScriptAnchor
msoShapeTypeMixed
msoTable
msoTextBox
msoTextEffect

#### **ShapeRange.Visible**

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。

**语法**

**express.Visible**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**：指定对象或对象格式是可见的。     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为当前演示文稿中第一张幻灯片的第三个形状设置阴影的水平偏移量和垂直偏移量，使其向右偏移 5 磅并向上偏移 3 磅。如果该形状原来没有阴影，本例将会为其添上阴影*/ function test(){ 	let myDocument = Application.ActivePresentation.Slides.Item(1) 	let myShadow = myDocument.Shapes.Item(3).Shadow 		myShadow.Visible = msoTrue 		myShadow.OffsetX = 5 		myShadow.OffsetY = -3 }` |

#### **ShapeRange.Width**

以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Width**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所有水平可用空间。要使以下示例执行，只能有两个打开的演示窗口*/ function test(){ 	//pplicatioAn.Windows.Arrange(ppArrangeTiled) 	let sngHeight = Application.ActivePresentation.Windows.Item(1).Height                     // available height 	let sngWidth = Application.ActivePresentation.Windows.Item(1).Width + Windows.Item(2).Width    // available width 	let win1 = Application.ActivePresentation.Windows.Item(1) 		win1.Width = sngWidth 		win1.Height = sngHeight / 2 		win1.Left = 0 	let win2 = Application.ActivePresentation.Windows.Item(2) 		win2.Width = sngWidth 		win2.Height = sngHeight / 2 		win2.Top = sngHeight / 2 		win2.Left = 0 }  /*以下示例将指定表中第一列的宽度设置为 80 磅*/ ActivePresentation.Slides.Item(2).Shapes.Item(2).Table.Columns.Item(1).Width = 80` |

#### **ShapeRange.ZOrderPosition**

返回指定的形状在 z-顺序中的位置。只读。Long 类型。

**语法**

**express.ZOrderPosition**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

Shapes(1) 返回 z-顺序中的最后一个形状，Shapes(Shapes.Count) 返回 z-顺序中的第一个形状。
若要设置形状在 z-顺序中的位置，请用 ZOrder 方法。
形状在 z-顺序中的位置与 Shapes 集合中形状的索引号相对应。例如，如果幻灯片上有四个形状，表达式 myDocument.Shapes(1) 在 z-顺序的后面返回该形状，表达式 myDocument.Shapes(4) 在 z-顺序的前面返回该形状。
如向一个集合添加新图形，其默认的添加位置是 z-顺序中的前面位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加椭圆，如果幻灯片中有至少一个其他形状，则将该椭圆放置在 z-顺序中倒数第二的位置上*/ function test(){ 	let myDocument = Application.ActivePresentation.Slides.Item(1) 	let myShape =  	 myDocument.Shapes.AddShape(msoShapeOval,  100, 100, 100, 300)  	 while(myShape.ZOrderPosition > 2) { 		   myShape.ZOrder(msoSendBackward) 		} }` |

适用环境：web

适用平台：windows/linux