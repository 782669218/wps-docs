#### **ShapeRange**



代表一个形状范围，即某个文档中的一组形状。一个形状范围可以只包含一个形状，也可以包含该文档中的所有形状。

**方法**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Align**                      | 对齐指定形状范围中的形状。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Apply**                      | 应用于使用 **PickUp** 方法复制的特定图形格式。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanvasCropBottom**           | 从绘图画布的底部裁剪一定百分比的绘图画布高度。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanvasCropLeft**             | 从绘图画布左侧裁剪一定百分比的绘图画布宽度。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanvasCropRight**            | 从绘图画布右侧裁剪一定百分比的绘图画布宽度。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanvasCropTop**              | 从绘图画布顶部裁剪一定百分比的绘图画布高度。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertToInlineShape**       | 将文档绘图层的指定形状转换为文字层的内嵌形状。只能转换代表图片、OLE 对象或 ActiveX 控件的形状。返回[InlineShape](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/InlineShape/InlineShape%20.htm#jsObject_InlineShape)**值** |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**                     | 删除指定区域的形状。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Distribute**                 | 在指定的形状范围内均匀分布形状。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Duplicate**                  | 创建指定的 **ShapeRange** 对象的副本，以标准的偏移将新图形区域从原图形添加至 **Shapes** 集合，然后返回 **Shape** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Flip**                       | 水平或垂直翻转一个图形。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Group**                      | 组合指定区域中的图形并将组合图形作为单个 **Shape** 对象返回。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementLeft**              | 将指定形状水平移动指定的磅数。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementRotation**          | 使指定的形状绕 Z 轴旋转指定的角度。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementTop**               | 以指定磅数垂直移动指定形状。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**                       | 返回集合中的单个 **Shape** 对象。返回Shape值                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PickUp**                     | 复制指定形状的格式。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScaleHeight**                | 按指定的比例缩放形状范围的高度。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScaleWidth**                 | 按指定比例调整形状的宽度。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**                     | 选择指定的形状范围。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetShapesDefaultProperties** | 将文档中默认形状的格式应用于指定的形状范围。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Ungroup**                    | 取消指定形状范围中所有组合形状的组合，分解指定形状或形状范围中图片和 OLE 对象的组合，将取消组合后的形状以单个 **ShapeRange** 对象的形式返回。返回ShapeRange值 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ZOrder**                     | 将集合中指定的形状区域移动到其他形状的前面或后面（也就是说，更改形状区域在 Z 顺序中的位置）。 |

**属性**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Adjustments**                | 返回一个 **Adjustments** 对象, 该对象包含所有对指定 **ShapeRange** 对象（代表自选图形或艺术字）进行调整操作的调整值。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlternativeText**            | 返回或设置与网页的图形相关联的可选文字。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Anchor**                     | 返回一个 **Range** 对象，该对象代表指定图形区域的锁定范围。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                | 返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoShapeType**              | 返回或设置指定的 **ShapeRange** 对象的图形类型，该对象不是代表线条或任意多边形，而是代表自选图形。**MsoAutoShapeType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundStyle**            | 设置或返回指定形状范围中形状的背景样式。可读/写 [MsoBackgroundStyleIndex](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoBackgroundStyleIndex%20%E6%9E%9A%E4%B8%BE.html)。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Callout**                    | 返回**CalloutFormat**对象，该对象包含指定图形的标注格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CanvasItems**                | 返回一个 **CanvasShapes** 对象，该对象代表绘图画布上图形的集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Child**                      | 如果图形是子图形或位于图形区域的所有图形都是同一父图形的子图形，则该属性值为 **True**。**MsoTriState** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**                      | 返回一个 **Long** 类型的值，该值代表集合中图形的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                    | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fill**                       | 返回一个 **FillFormat**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/FillFormat/FillFormat%20.htm#jsObject_FillFormat)对象，该对象包含指定图形的填充格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Glow**                       | 返回一个 **GlowFormat** 对象，该对象代表形状区域的发光格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GroupItems**                 | 返回一个 **GroupShapes** 对象，该对象代表指定图形组中的单个图形。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**                     | 返回或设置指定图形区域的高度。**Single** 类型，可读写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeightRelative**             | 返回或设置一个 **Single** 类型的值，该值代表将形状区域大小调整到的目标形状的百分比。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalFlip**             | 表示该形状范围已进行水平翻转。只读 **MsoTriState** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyperlink**                  | 返回一个 **Hyperlink** 对象，该对象代表与指定 **ShapeRange** 对象相关联的超链接。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ID**                         | 返回形状范围的标识类型。只读 **Long** 类型。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LayoutInCell**               | 返回一个 **Long** 类型的值，该值代表表格中的形状是显示在表格内部还是表格外部。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                       | 返回或设置一个 **Single** 类型的值，该值代表指定形状范围的水平位置，以磅为单位。也可以是任何有效的 [WdShapePosition ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapePosition%20%E6%9E%9A%E4%B8%BE.html)常量。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftRelative**               | 返回或设置一个 **Single** 类型的值，该值代表形状区域左侧的相对位置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Line**                       | 返回一个 **LineFormat** 对象，该对象包含指定形状范围的线条格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LockAnchor**                 | 如果指定 **ShapeRange** 对象的锁定标记锁定到锁定范围，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LockAspectRatio**            | 如果在调整指定形状的大小时保留其最初比例，则该属性值为 **MsoTrue**；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 **MsoFalse**。可读写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                       | 返回或设置指定对象的名称。**String** 类型，可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Nodes**                      | 返回一个 **ShapeNodes** 集合，该集合代表指定形状的几何描述。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                     | 返回一个 **Object** 类型值，该值代表指定 **ShapeRange** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentGroup**                | 返回一个 **Shape** 对象，该对象代表形状范围的通用父形状。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureFormat**              | 返回一个 **PictureFormat** 对象，该对象包含指定形状范围的图片格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reflection**                 | 返回一个**ReflectionFormat**对象，该对象代表形状区域的反射格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelativeHorizontalPosition** | 指定形状范围的相对水平位置。可读写 [WdRelativeHorizontalPosition ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeHorizontalPosition%20%E6%9E%9A%E4%B8%BE.html)类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelativeHorizontalSize**     | 返回或设置一个 [WdRelativeHorizontalSize ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeHorizontalSize%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表形状区域相对的对象。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelativeVerticalPosition**   | 指定形状范围的相对垂直位置。可读写 [WdRelativeHorizontalPosition ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeHorizontalPosition%20%E6%9E%9A%E4%B8%BE.html)类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelativeVerticalSize**       | 返回或设置一个 **WdRelativeVerticalSize**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeVerticalSize%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表形状区域相对的对象。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rotation**                   | 返回或设置指定形状绕 Z 轴旋转的度数。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**                     | 返回一个 **ShadowFormat** 对象，该对象代表指定形状的阴影格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShapeStyle**                 | 设置或返回指定形状范围中形状的形状样式。可读/写 MsoShapeStyleIndex。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SoftEdge**                   | 返回一个**SoftEdgeFormat** 对象，该对象代表形状区域的软边缘格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextEffect**                 | 返回一个 **TextEffectFormat** 对象，该对象包含指定形状的文本效果格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame**                  | 返回一个 **TextFrame** 对象，该对象包含指定形状范围的文字。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame2**                 | 返回一个 **TextFrame2** 对象，包含指定形状区域的文本。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThreeD**                     | 返回一个 **ThreeDFormat** 对象，该对象包含指定形状范围的三维格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Title**                      | 返回或设置 **String** 类型值，该值包含指定形状范围中形状的标题。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                        | 返回或设置指定形状或形状范围的垂直位置（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TopRelative**                | 返回或设置一个 **Single** 类型的值，该值代表形状区域顶部的相对位置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                       | 返回形状类型。只读 **MsoShapeType** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalFlip**               | 如果指定形状围绕垂直轴进行翻转，则该属性值为 **True**。**MsoTriState** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Vertices**                   | 该属性以一系列坐标对的形式返回指定任意多边形图形顶点（和贝赛尔曲线的控点）的坐标。可将该属性返回的数组用作 **AddCurve** 或 **AddPolyLine** 方法的参数。只读 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**                    | 如果指定对象或应用于该对象的格式是可见的，则该属性值为 **True**。**MsoTriState** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**                      | 返回或设置范围内形状的宽度（以磅为单位）。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WidthRelative**              | 返回或设置一个 **Single** 类型的值，该值代表形状区域的相对宽度。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WrapFormat**                 | 返回一个 **WrapFormat** 对象，该对象包含在指定的形状范围四周文字环绕的属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ZOrderPosition**             | 返回一个 **Long** 类型的值，该值代表指定的形状在 Z 顺序中的位置。只读。 |

**成员方法**

#### **ShapeRange.Align**

对齐指定形状范围中的形状。

**语法**

**express.Align(Align, RelativeTo)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**    | **说明**                                                     |
| ------------ | ------------- | --------------- | ------------------------------------------------------------ |
| *Align*      | 必选          | **MsoAlignCmd** | 指定特定形状范围中形状的对齐方式。                           |
| *RelativeTo* | 必选          | **Long**        | 如果该参数值为 True，则相对于文档边缘对齐形状。如果该参数值为 False，则相对于彼此对齐形状。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中选定范围内所有形状的左边缘与该范围最左边的形状的左边缘对齐。*/ function test() {     let myShapeRange = Application.Selection.ShapeRange     myShapeRange.Align(msoAlignLefts, false) }  ` |

#### **ShapeRange.Apply**

应用于使用 **PickUp** 方法复制的特定图形格式。

**语法**

**express.Apply()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

如果以前没有使用 **PickUp** 方法复制 **ShapeRange** 对象的格式，则使用**Apply**方法会导致错误。

#### **ShapeRange.CanvasCropBottom**

从绘图画布的底部裁剪一定百分比的绘图画布高度。

**语法**

**express.CanvasCropBottom(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 裁剪后所需保留的绘图画布高度的百分比数值。输入 0.9 作为增量从底部裁剪绘图画布高度的百分之十。输入 0.1 从底部裁剪绘图画布高度的百分之九十。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*假定活动文档中的第一个图形为绘图画布的情况下，则本示例从活动文档中第一块绘图画布的底部裁剪绘图画布高度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在文档中添加绘图画布。*/ function test() {     let shpCanvas = Application.ActiveDocument.Shapes.Item(1)     shpCanvas.CanvasCropBottom(0.75) }` |

#### **ShapeRange.CanvasCropLeft**

从绘图画布左侧裁剪一定百分比的绘图画布宽度。

**语法**

**express.CanvasCropLeft(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 裁剪绘图画布后所需保留的绘图画布宽度的百分比数值。输入 0.9 作为增量从左侧裁剪绘图画布宽度的百分之十。输入 0.1 从左侧裁剪绘图画布宽度的百分之九十。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从活动文档中第一块绘图画布的左侧裁剪绘图画布宽度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在文档中添加绘图画布。*/ function test() {     let shpCanvas = Application.ActiveDocument.Shapes.Item(1)     shpCanvas.CanvasCropLeft(0.75) }` |

#### **ShapeRange.CanvasCropRight**

从绘图画布右侧裁剪一定百分比的绘图画布宽度。

**语法**

**express.CanvasCropRight(Incremet)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Incremet* | 必选          | **Single**   | 裁剪绘图画布后所需保留的绘图画布宽度的百分比数值。输入 0.9 作为增量从右侧裁剪绘图画布宽度的百分之十。输入 0.1 从右侧裁剪绘图画布宽度的百分之九十。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从右侧裁剪活动文档中第一块绘图画布宽度的百分之二十五。如果不是这种情况，需要使用 AddCanvas 方法在文档中添加绘图画布。*/ function test(){     let shpCanvas = Application.ActiveDocument.Shapes.Item(1)     shpCanvas.CanvasCropRight(0.75) }` |

#### **ShapeRange.CanvasCropTop**

从绘图画布顶部裁剪一定百分比的绘图画布高度。

**语法**

**express.CanvasCropTop(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 裁剪绘图画布后所需保留的绘图画布高度的百分比数值。输入 0.9 作为增量从顶部裁剪绘图画布高度的百分之十。输入 0.1 从顶部裁剪绘图画布高度的百分之九十。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*假定活动文档中的第一个图形为绘图画布的情况下，本示例从活动文档中第一块绘图画布的顶部裁剪绘图画布高度的百分之二十五。如果不是这种情况，则需要使用 AddCanvas 方法在活动文档中添加绘图画布。*/ function test() {     let shpCanvas = Application.ActiveDocument.Shapes.Item(1)     shpCanvas.CanvasCropTop(0.75) }` |

#### **ShapeRange.ConvertToInlineShape**

将文档绘图层的指定形状转换为文字层的内嵌形状。只能转换代表图片、OLE 对象或 ActiveX 控件的形状。

返回[InlineShape](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/InlineShape/InlineShape%20.htm#jsObject_InlineShape)**值**

**语法**

**express.ConvertToInlineShape()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

支持附加文字的形状不能转换为内嵌形状。对于这样的形状，请使用 **ConvertToFrame** 方法。

如果对包含多个形状的 **ShapeRange** 对象使用此方法，将导致出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 MyDoc.doc 中的每张图片转换为内嵌形状。*/  function test() {   for(let i = 1; i <= Application.Documents.Item("MyDoc.doc").Shapes.Count; i++){     if(Application.Documents.Item("MyDoc.doc").Shapes.Item(i).Type == msoPicture){         Application.Documents.Item("MyDoc.doc").Shapes.Item(i).ConvertToInlineShape()     }   } }` |

#### **ShapeRange.Delete**

删除指定区域的形状。

**语法**

**express.Delete()**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Distribute**

在指定的形状范围内均匀分布形状。

**语法**

**express.Distribute(Distribute, RelativeTo)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**         | **说明**                                                     |
| ------------ | ------------- | -------------------- | ------------------------------------------------------------ |
| *Distribute* | 必选          | **MsoDistributeCmd** | 指定图形是横向分布还是纵向分布。                             |
| *RelativeTo* | 必选          | **Long**             | 如果为 True，则将图形均布在页面的整个横向或者纵向空间上。如果为 False，则图形在其原来占有的横向或者纵向空间内分布。 |

**说明**

可选择纵向分布或横向分布，也可选择将各图形分布于整个页面还是仅限于原先占据的空间。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例可实现的功能为：定义一个图形区域，该区域包含活动文档中所有的自选图形，再将各图形水平分布在此区域内。*/ function test() {   let docShapes = Application.ActiveDocument.Shapes   let numShapes = docShapes.Count   let autoShpArray = []   let numAutoShapes   let asRange   if(numShapes > 1){      numAutoShapes = 0     for(let i = 1; i <= numShapes; i++) {         if(docShapes.Item(i).Type == msoAutoShape) {             numAutoShapes++             autoShpArray[numAutoShapes] = docShapes.Item(i).Name         }     }     if(numAutoShapes > 1) {         asRange = docShapes.Range(autoShpArray)               asRange.Distribute(msoDistributeHorizontally, false)     }   } }` |

#### **ShapeRange.Duplicate**

创建指定的 **ShapeRange** 对象的副本，以标准的偏移将新图形区域从原图形添加至 **Shapes** 集合，然后返回 **Shape** 对象。

**语法**

**express.Duplicate()**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建活动文档第一个图形的副本，然后改变新图形的填充格式。*/ function test() {     let newShape = Application.ActiveDocument.Shapes.Item(1).Duplicate()     newShape.Fill.PresetGradient(msoGradientVertical, 1, msoGradientGold) }` |

#### **ShapeRange.Flip**

水平或垂直翻转一个图形。

**语法**

**express.Flip(FlipCmd)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**   | **说明**   |
| --------- | ------------- | -------------- | ---------- |
| *FlipCmd* | 必选          | **MsoFlipCmd** | 翻转方向。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先向活动文档添加一个三角形，然后复制此三角形，再垂直翻转复制的三角形并将其设为红色。*/ function test(){     let shapes = Application.ActiveDocument.Shapes.AddShape(msoShapeRightTriangle, 150, 150, 50, 50).Duplicate()     shapes.Fill.ForeColor.RGB = (255, 0, 0)     shapes.Flip(msoFlipVertical) }` |

#### **ShapeRange.Group**

组合指定区域中的图形并将组合图形作为单个 **Shape** 对象返回。

**语法**

**express.Group()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

由于一组形状作为单个形状处理，所以创建和分解形状组将改变 **Shapes** 集合中的项目数，而且由于影响集合中的项目，还会改变部分项目的索引号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加两个图形，并组合这两个新图形，接着设置该组的填充格式并对该组进行旋转，然后将该组置于绘图层的后面。*/ function test() {     let myDocument = Application.ActiveDocument.Shapes     myDocument.AddShape(msoShapeCan, 50, 10, 100, 200).Name = "shpOne"     myDocument.AddShape(msoShapeCube, 150, 250, 100, 200).Name = "shpTwo"     let newmyDocument = myDocument.Range(Array("shpOne", "shpTwo")).Group()     newmyDocument.Fill.PresetTextured(msoTextureBlueTissuePaper)     newmyDocument.Rotation = 45     newmyDocument.ZOrder(msoSendToBack) }` |

#### **ShapeRange.IncrementLeft**

将指定形状水平移动指定的磅数。

**语法**

**express.IncrementLeft(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定形状水平移动的距离，以磅为单位。为正值时将形状右移；为负值时将形状左移。 |

#### **ShapeRange.IncrementRotation**

使指定的形状绕 Z 轴旋转指定的角度。

**语法**

**express.IncrementRotation(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定形状的水平旋转量，以度为单位。为正值时顺时针旋转形状，为负值时逆时针旋转形状。 |

**说明**

使用 **Rotation** 属性可以设置形状的绝对旋转角度。如果要将三维形状绕 X 轴或 Y 轴旋转，请使用 **ThreeDFormat** 的 **IncrementRotationX** 或 **IncrementRotationY** 方法。

#### **ShapeRange.IncrementTop**

以指定磅数垂直移动指定形状。

**语法**

**express.IncrementTop(Increment)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定对象的垂直移动距离，以磅为单位。为正值时将形状下移；为负值时将形状上移。 |

#### **ShapeRange.Item**

返回集合中的单个 **Shape** 对象。返回Shape值

**语法**

**express.Item(Index)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

#### **ShapeRange.PickUp**

复制指定形状的格式。

**语法**

**express.PickUp()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

用 [Apply ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.Apply)方法可将复制的格式应用于另一个形状。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例首先复制 myDocument 上第一个形状的格式，然后将复制的形状格式应用于第二个形状。*/ Application.ActiveDocument.Shapes.Item(1).PickUp() Application.ActiveDocument.Shapes.Item(2).Apply()` |

#### **ShapeRange.ScaleHeight**

按指定的比例缩放形状范围的高度。

**语法**

**express.ScaleHeight(Factor, RelativeToOriginalSize, Scale)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**     | **说明**                                                     |
| ------------------------ | ------------- | ---------------- | ------------------------------------------------------------ |
| *Factor*                 | 必选          | **Single**       | 指定形状调整后的高度与当前或原始高度的比例。例如，要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选          | **MsoTriState**  | 如果该参数值为 True，则相对于原始大小缩放形状。如果该参数值为 False，则相对于当前大小缩放形状。仅当指定的形状为图片或 OLE 对象时，才能将此参数指定为 True。 |
| *Scale*                  | 可选          | **MsoScaleFrom** | 在缩放形状时，形状中位置不变的部分。                         |

**说明**

对于图片和 OLE 对象，您可以说明是相对于原始大小还是相对于当前大小缩放形状。对于图片和 OLE 对象以外的形状总是相对于当前高度缩放。

#### **ShapeRange.ScaleWidth**

按指定比例调整形状的宽度。

**语法**

**express.ScaleWidth(Factor, RelativeToOriginalSize, Scale)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**     | **说明**                                                     |
| ------------------------ | ------------- | ---------------- | ------------------------------------------------------------ |
| *Factor*                 | 必选          | **Single**       | 指定形状调整后的宽度与当前或原始宽度的比例。例如，要将一个矩形放大百分之五十，请将此参数指定为 1.5。 |
| *RelativeToOriginalSize* | 必选          | **MsoTriState**  | 如果该参数值为 True，则相对于原始大小缩放形状。如果该参数值为 False，则相对于当前大小缩放形状。仅当指定的形状为图片或 OLE 对象时，才能将此参数指定为 True。 |
| *Scale*                  | 可选          | **MsoScaleFrom** | 在缩放形状时，形状中位置不变的部分。                         |

**说明**

对于图片和 OLE 对象，您可以说明是相对于原始大小还是相对于当前大小缩放形状的范围。图片和 OLE 对象以外的形状总是相对于当前宽度缩放。

#### **ShapeRange.Select**

选择指定的形状范围。

**语法**

**express.Select(Replace)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | 在添加形状时，如果该参数值为 True，则替换所选内容。如果该参数值为 False，则将新形状添加到所选内容。 |

#### **ShapeRange.SetShapesDefaultProperties**

将文档中默认形状的格式应用于指定的形状范围。

**语法**

**express.SetShapesDefaultProperties()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

新形状将继承默认形状的许多属性。

#### **ShapeRange.Ungroup**

取消指定形状范围中所有组合形状的组合，分解指定形状或形状范围中图片和 OLE 对象的组合，将取消组合后的形状以单个 **ShapeRange** 对象的形式返回。

返回ShapeRange值

**语法**

**express.Ungroup()**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

由于将组合形状作为单个对象来处理，对形状进行组合和取消组合将会改变 **Shapes** 集合中项目的数量，并且改变集合中取消组合的形状后面的项目的索引号。

#### **ShapeRange.ZOrder**

将集合中指定的形状区域移动到其他形状的前面或后面（也就是说，更改形状区域在 Z 顺序中的位置）。

**语法**

**express.ZOrder(ZOrderCmd)**

*express*   一个代表 **ShapeRange** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**     | **说明**                                         |
| ----------- | ------------- | ---------------- | ------------------------------------------------ |
| *ZOrderCmd* | 必选          | **MsoZOrderCmd** | 指定相对于其他形状将指定的形状区域移动到的位置。 |

**说明**

使用 [ZOrderPosition ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.ZOrderPosition)属性确定形状区域在 Z 顺序中的当前位置。

**成员属性**

#### **ShapeRange.Adjustments**

返回一个 **Adjustments** 对象, 该对象包含所有对指定 **ShapeRange** 对象（代表自选图形或艺术字）进行调整操作的调整值。只读。

**语法**

**express.Adjustments**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.AlternativeText**

返回或设置与网页的图形相关联的可选文字。**String** 类型，可读写。

**语法**

**express.AlternativeText**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：为活动窗口中选定的图形设定可选文字。选定图形是一幅野鸭的图*/ Application.ActiveWindow.Selection.ShapeRange.AlternativeText = "This is a mallard duck."` |

#### **ShapeRange.Anchor**

返回一个 **Range** 对象，该对象代表指定图形区域的锁定范围。只读。

**语法**

**express.Anchor**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

如果对包含多个图形的 **ShapeRange** 对象使用该属性，将导致出错。

所有 **Shape** 对象都锁定于某一个文字区域，但可将其置于锁定标记所在页的任意位置。如果创建图形时指定了锁定范围，则锁定标记位于包含该锁定范围的第一段落的段首。如果没有指定锁定范围，则将自动选择锁定范围，图形参照页面的左边框和上边框进行定位。

图形总是与锁定标记处在同一页上。如果图形的 **LockAnchor** 属性为 **True**，则不能在页面上拖动锁定标记。

#### **ShapeRange.Application**

返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **ShapeRange.AutoShapeType**

返回或设置指定的 **ShapeRange** 对象的图形类型，该对象不是代表线条或任意多边形，而是代表自选图形。**MsoAutoShapeType** 类型，可读写。

**语法**

**express.AutoShapeType**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

改变一个形状的类型时，该形状保留其大小、颜色和其他属性。

#### **ShapeRange.BackgroundStyle**

设置或返回指定形状范围中形状的背景样式。可读/写 [MsoBackgroundStyleIndex](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoBackgroundStyleIndex%20%E6%9E%9A%E4%B8%BE.html)。

**语法**

**express.BackgroundStyle**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Callout**

返回**CalloutFormat**对象，该对象包含指定图形的标注格式属性。只读。

**语法**

**express.Callout**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

该属性应用于代表标注的 **ShapeRange** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*该示例在 myDocument 中添加一个椭圆和一个指向椭圆的标注。标注文本不带有边框，但有一条强调线将文本和标注线隔开。*/ function test() {     let myShapes = Application.ActiveDocument.Shapes     myShapes.AddShape(msoShapeOval, 180, 200, 280, 130)      let newCallout = myShapes.AddCallout(msoCalloutTwo, 420, 170, 170, 40)     newCallout.TextFrame.TextRange.Text = "My oval"     newCallout.Callout.Accent = true     newCallout.Callout.Border = false }` |

#### **ShapeRange.CanvasItems**

返回一个 **CanvasShapes** 对象，该对象代表绘图画布上图形的集合。

**语法**

**express.CanvasItems**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Child**

如果图形是子图形或位于图形区域的所有图形都是同一父图形的子图形，则该属性值为 **True**。**MsoTriState** 类型，只读。

**语法**

**express.Child**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Count**

返回一个 **Long** 类型的值，该值代表集合中图形的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **ShapeRange.Fill**

返回一个 **FillFormat**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/FillFormat/FillFormat%20.htm#jsObject_FillFormat)对象，该对象包含指定图形的填充格式属性。只读。

**语法**

**express.Fill**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Glow**

返回一个 **GlowFormat** 对象，该对象代表形状区域的发光格式。只读。

**语法**

**express.Glow**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.GroupItems**

返回一个 **GroupShapes** 对象，该对象代表指定图形组中的单个图形。只读。

**语法**

**express.GroupItems**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

该属性应用于代表组合图形的 **ShapeRange** 对象。使用 **GroupShapes** 对象的 **Item** 方法可从图形组中返回单个图形。

#### **ShapeRange.Height**

返回或设置指定图形区域的高度。**Single** 类型，可读写。

**语法**

**express.Height**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.HeightRelative**

返回或设置一个 **Single** 类型的值，该值代表将形状区域大小调整到的目标形状的百分比。可读写。

**语法**

**express.HeightRelative**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

将此属性与 [RelativeVerticalSize ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.RelativeVerticalSize)属性一起使用。当设置为 **wdShapeSizeRelativeNone**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapeSizeRelative%20%E6%9E%9A%E4%B8%BE.html)(-999999)（参见 **WdShapeSizeRelative**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapeSizeRelative%20%E6%9E%9A%E4%B8%BE.html)枚举）时，应忽略此属性，因为形状不能使用百分比大小。高度是由 [Height ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.Height)属性单独确定的。

#### **ShapeRange.HorizontalFlip**

表示该形状范围已进行水平翻转。只读 **MsoTriState** 类型。

**语法**

**express.HorizontalFlip**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Hyperlink**

返回一个 **Hyperlink** 对象，该对象代表与指定 **ShapeRange** 对象相关联的超链接。只读。

**语法**

**express.Hyperlink**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果不存在与指定形状区域相关联的超链接，则会发生错误。在此情况下，请将 Add 方法用于 Hyperlinks 集合来为指定形状区域添加超链接。以下示例说明了具体的操作方法。*/ Application.ActiveDocument.Hyperlinks.Add(Selection.ShapeRange.Item(1), "http://www.microsoft.com")` |

#### **ShapeRange.ID**

返回形状范围的标识类型。只读 **Long** 类型。

**语法**

**express.ID**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.LayoutInCell**

返回一个 **Long** 类型的值，该值代表表格中的形状是显示在表格内部还是表格外部。

**语法**

**express.LayoutInCell**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

**LayoutInCell** 属性对应于图片格式的**“高级版式”**对话框中的**“表格单元格中的版式”**选项。如果为 **True**，则表示指定的图片显示在表格内部。如果为 **False**，则表示指定的图片显示在表格外部。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释         |
| ------------------------------------------------------------ |
| 仅当 **WrapFormat** 对象的 **Type** 属性设置为除 **wdWrapTypeInline** 或 **wdWrapTypeNone** 之外的值时，对 **LayoutInCell** 属性的设置才会生效。 |

#### **ShapeRange.Left**

返回或设置一个 **Single** 类型的值，该值代表指定形状范围的水平位置，以磅为单位。也可以是任何有效的 [WdShapePosition ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapePosition%20%E6%9E%9A%E4%B8%BE.html)常量。可读写。

**语法**

**express.Left**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

形状的位置是根据形状边框的左上角到形状锁定标记的相对距离进行计算的。**RelativeHorizontalPosition** 属性控制锁定标记是沿字符、文本栏、页边距还是页面边缘放置。

对于包含多个形状的 **ShapeRange** 对象，**Left** 属性设置每个形状的水平位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一、第二个形状的水平位置设置为距离文本栏的左边缘 1 英寸。*/  function test() {   let shapes = Application.ActiveDocument.Shapes.Range([1,2])   shapes.RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn   shapes.Left = InchesToPoints(1) }` |

#### **ShapeRange.LeftRelative**

返回或设置一个 **Single** 类型的值，该值代表形状区域左侧的相对位置。可读写。

**语法**

**express.LeftRelative**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

将此属性与 [RelativeHorizontalPosition ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeHorizontalPosition%20%E6%9E%9A%E4%B8%BE.html)属性一起使用。当设置为 **wdShapePositionRelativeNone**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/wdShapePositionRelativeNone)(-999999)（参见 [WdShapePositionRelative ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapePositionRelative%20%E6%9E%9A%E4%B8%BE.html)枚举）时，应忽略此属性，因为形状不能使用百分比放置。水平位置是由 [Left ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.Left)属性单独确定的。

#### **ShapeRange.Line**

返回一个 **LineFormat** 对象，该对象包含指定形状范围的线条格式属性。只读。

**语法**

**express.Line**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

对于线条来说，**LineFormat** 对象代表线条本身；而对于带有边框的形状范围来说，**LineFormat** 对象代表边框。

#### **ShapeRange.LockAnchor**

如果指定 **ShapeRange** 对象的锁定标记锁定到锁定范围，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.LockAnchor**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

如果形状范围有一个锁定的标记，则不能通过拖动来移动形状的锁定标记。锁定标记不会随形状移动而移动。

虽然 **ShapeRange** 对象已锁定到某一文字范围，但是您可以将该对象放在页面上的任何位置。如果形状范围锁定到包含锁定范围的第一段的开头，则形状总是与其锁定标记位于同一页上。

#### **ShapeRange.LockAspectRatio**

如果在调整指定形状的大小时保留其最初比例，则该属性值为 **MsoTrue**；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 **MsoFalse**。可读写 **MsoTriState** 类型。

**语法**

**express.LockAspectRatio**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Name**

返回或设置指定对象的名称。**String** 类型，可读写。

**语法**

**express.Name**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Nodes**

返回一个 **ShapeNodes** 集合，该集合代表指定形状的几何描述。

**语法**

**express.Nodes**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Parent**

返回一个 **Object** 类型值，该值代表指定 **ShapeRange** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.ParentGroup**

返回一个 **Shape** 对象，该对象代表形状范围的通用父形状。

**语法**

**express.ParentGroup**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.PictureFormat**

返回一个 **PictureFormat** 对象，该对象包含指定形状范围的图片格式属性。只读。

**语法**

**express.PictureFormat**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

适用于代表图片或 OLE 对象的 **ShapeRange** 对象。

#### **ShapeRange.Reflection**

返回一个**ReflectionFormat**对象，该对象代表形状区域的反射格式。只读。

**语法**

**express.Reflection**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.RelativeHorizontalPosition**

指定形状范围的相对水平位置。可读写 [WdRelativeHorizontalPosition ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeHorizontalPosition%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.RelativeHorizontalPosition**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例重新放置选定形状对象。*/   function test() {   let shaperange = Application.Selection.ShapeRange   shaperange.Left = InchesToPoints(0.6)   shaperange.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage   shaperange.Top = InchesToPoints(1)   shaperange.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph }` |

#### **ShapeRange.RelativeHorizontalSize**

返回或设置一个 [WdRelativeHorizontalSize ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeHorizontalSize%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表形状区域相对的对象。可读写。

**语法**

**express.RelativeHorizontalSize**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

此属性可与 [WidthRelative ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/WidthRelative)属性一起使用。

#### **ShapeRange.RelativeVerticalPosition**

指定形状范围的相对垂直位置。可读写 [WdRelativeHorizontalPosition ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeHorizontalPosition%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.RelativeVerticalPosition**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例重新放置选定形状对象。*/  function test() { let shaperange = Application.Selection.ShapeRange shaperange.Left = InchesToPoints(0.6) shaperange.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage shaperange.Top = InchesToPoints(1) shaperange.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph }` |

#### **ShapeRange.RelativeVerticalSize**

返回或设置一个 **WdRelativeVerticalSize**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeVerticalSize%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表形状区域相对的对象。可读写。

**语法**

**express.RelativeVerticalSize**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

此属性可与 [HeightRelative ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.HeightRelative)属性一起使用。

#### **ShapeRange.Rotation**

返回或设置指定形状绕 Z 轴旋转的度数。可读写 **Single** 类型。

**语法**

**express.Rotation**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

正值表示顺时针旋转；负值表示逆时针旋转。要设置三维形状绕 X 轴或 Y 轴的旋转，请使用**ThreeDFormat** 对象的 **RotationX** 属性或 **RotationY** 属性。

#### **ShapeRange.Shadow**

返回一个 **ShadowFormat** 对象，该对象代表指定形状的阴影格式。

**语法**

**express.Shadow**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.ShapeStyle**

设置或返回指定形状范围中形状的形状样式。可读/写 MsoShapeStyleIndex。

**语法**

**express.ShapeStyle**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.SoftEdge**

返回一个**SoftEdgeFormat** 对象，该对象代表形状区域的软边缘格式。只读。

**语法**

**express.SoftEdge**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.TextEffect**

返回一个 **TextEffectFormat** 对象，该对象包含指定形状的文本效果格式属性。只读。

**语法**

**express.TextEffect**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

适用于代表艺术字的 **ShapeRange** 对象。

#### **ShapeRange.TextFrame**

返回一个 **TextFrame** 对象，该对象包含指定形状范围的文字。

**语法**

**express.TextFrame**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.TextFrame2**

返回一个 **TextFrame2** 对象，包含指定形状区域的文本。只读。

**语法**

**express.TextFrame2**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.ThreeD**

返回一个 **ThreeDFormat** 对象，该对象包含指定形状范围的三维格式属性。只读。

**语法**

**express.ThreeD**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Title**

返回或设置 **String** 类型值，该值包含指定形状范围中形状的标题。可读/写。

**语法**

**express.Title**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

使用 **Title** 属性可为形状提供可选文字标题。该属性将标题文本添加到 WPS 2015 中**“设置形状格式”**对话框的**“可选文字”**窗格上的**“标题”**文本框中。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释         |
| ------------------------------------------------------------ |
| Web 浏览器在加载表格的过程中或表格丢失时显示可选文字。Web 搜索引擎利用可选文字帮助查找网页。可选文字也可用来帮助残障人士。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例向活动文档中第一个和第三个形状中添加可选文字标题。*/ Application.ActiveDocument.Shapes.Range([1, 3]).Title = "Part of a shape array."` |

#### **ShapeRange.Top**

返回或设置指定形状或形状范围的垂直位置（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.Top**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

形状的位置是根据形状边框的左上角到形状锁定标记的相对距离进行计算的。**RelativeVerticalPosition** 属性控制形状的锁定标记是沿行、段落、页边距还是页面边缘放置。

对于包含多个形状的 **ShapeRange** 对象，**Top** 属性设置每个形状的垂直位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个和第二个形状的垂直位置设置为距页面顶部 1 英寸。*/  function test() {   let shapes = Application.ActiveDocument.Shapes.Range(Array(1, 2))   shapes.RelativeVerticalPosition = wdRelativeVerticalPositionPage   shapes.Top = InchesToPoints(1) }` |

#### **ShapeRange.TopRelative**

返回或设置一个 **Single** 类型的值，该值代表形状区域顶部的相对位置。可读写。

**语法**

**express.TopRelative**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

将此属性与 [RelativeVerticalPosition ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeVerticalPosition%20%E6%9E%9A%E4%B8%BE.html)属性一起使用。当设置为 [wdShapePositionRelativeNone](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapePositionRelative%20%E6%9E%9A%E4%B8%BE.html)[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapePositionRelative%20%E6%9E%9A%E4%B8%BE.html)(-999999)（参见 [WdShapePositionRelative ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapePositionRelative%20%E6%9E%9A%E4%B8%BE.html)枚举）时，应忽略此属性，因为形状不能使用百分比放置。垂直位置是由 [Top ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.Top)属性单独确定的。

#### **ShapeRange.Type**

返回形状类型。只读 **MsoShapeType** 类型。

**语法**

**express.Type**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.VerticalFlip**

如果指定形状围绕垂直轴进行翻转，则该属性值为 **True**。**MsoTriState** 类型，只读。

**语法**

**express.VerticalFlip**

*express*   一个代表 **ShapeRange** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 myDocument 中所有进行过水平翻转或垂直翻转的形状还原至初始状态。*/  function test() {   let s   for(let i =1; i <= Application.ActiveDocument.Range().ShapeRange.Count; i++) {       s = Application.ActiveDocument.Range().ShapeRange.Item(i)       if(s.HorizontalFlip) {           s.Flip(msoFlipHorizontal)       }       if(s.VerticalFlip) {            s.Flip(msoFlipVertical)       }   } }` |

#### **ShapeRange.Vertices**

该属性以一系列坐标对的形式返回指定任意多边形图形顶点（和贝赛尔曲线的控点）的坐标。可将该属性返回的数组用作 **AddCurve** 或 **AddPolyLine** 方法的参数。只读 **Variant** 类型。

**语法**

**express.Vertices**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

下表显示 **Vertices** 属性如何将数组 *vertArray()* 的值与三角形的顶点坐标相关联。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `vertArray(1, 1)vertArray(1, 2)vertArray(2, 1)vertArray(2, 2)vertArray(3, 1)vertArray(3, 2)` |

#### **ShapeRange.Visible**

如果指定对象或应用于该对象的格式是可见的，则该属性值为 **True**。**MsoTriState** 类型，可读写。

**语法**

**express.Visible**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.Width**

返回或设置范围内形状的宽度（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Width**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.WidthRelative**

返回或设置一个 **Single** 类型的值，该值代表形状区域的相对宽度。可读写。

**语法**

**express.WidthRelative**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

将此属性与 [RelativeHorizontalSize ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdRelativeHorizontalSize%20%E6%9E%9A%E4%B8%BE.html)属性一起使用。当设置为 [wdShapeSizeRelativeNone](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapeSizeRelative%20%E6%9E%9A%E4%B8%BE.html)[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapeSizeRelative%20%E6%9E%9A%E4%B8%BE.html)(-999999)（参见 [WdShapeSizeRelative ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdShapeSizeRelative%20%E6%9E%9A%E4%B8%BE.html)枚举）时，应忽略此属性，因为形状不能使用百分比大小。宽度是由 [Width ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#ShapeRange.Width)属性单独确定的。

#### **ShapeRange.WrapFormat**

返回一个 **WrapFormat** 对象，该对象包含在指定的形状范围四周文字环绕的属性。只读。

**语法**

**express.WrapFormat**

*express*   一个代表 **ShapeRange** 对象的变量。

#### **ShapeRange.ZOrderPosition**

返回一个 **Long** 类型的值，该值代表指定的形状在 Z 顺序中的位置。只读。

**语法**

**express.ZOrderPosition**

*express*   一个代表 **ShapeRange** 对象的变量。

**说明**

`Shapes(1)` 返回 Z 顺序中的最后一个形状，而 `Shapes(Shapes.Count)` 返回 Z 顺序中的第一个形状。该属性为只读。要设置形状在 Z 顺序中的位置，请使用 **ZOrder** 方法。

形状在 Z 顺序中的位置与 Shapes 集合中形状的索引号相对应。例如，如果在 myDocument 上有四个形状，则表达式 `myDocument.Shapes(1)` 返回 Z 顺序中的最后一个形状，而表达式 `myDocument.Shapes(4)` 返回 Z 顺序中的第一个形状。

无论何时将新形状添加到集合中，默认情况下该形状都将添加到 Z 顺序的最前端。

适用环境：web

适用平台：windows/linux