#### **FillFormat**



代表形状的填充格式。形状可以有单色、渐变、纹理、图案、图片或半透明填充。

**说明**

使用 **Fill** 属性可返回 **FillFormat** 对象。以下示例向活动文档添加一个矩形，然后设置该矩形填充的渐变和颜色。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let shp = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80); 	let fill = shp.Fill; 	fill.ForeColor.RGB = 0x00FFFF 	fill.OneColorGradient(msoGradientHorizontal, 1, 1); }` |

**方法**

|                                                              | 名称                 | 说明                         |
| ------------------------------------------------------------ | -------------------- | ---------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OneColorGradient** | 将指定填充设置为单色渐变。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Patterned**        | 将指定填充设置为一种图案     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Solid**            | 将指定填充设置为统一的颜色。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TwoColorGradient** | 将指定填充设为双色渐变。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UserPicture**      | 用图像填充指定的形状。       |

**属性**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackColor**         | 返回或设置一个 **ColorFormat** 对象，该对象代表填充的背景色，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ForeColor**         | 返回或设置一个 **ColorFormat** 对象，该对象代表填充的前景色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GradientAngle**     | 返回或设置指定填充格式的渐变填充的角度。可读/写。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GradientColorType** | 返回指定填充的渐变颜色的样式。只读 **MsoGradientColorType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GradientDegree**    | 返回值表明单色渐变填充的深浅程度。**Number** 类型，只读。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Pattern**           | 返回或设置一个 **MsoPatternType** 常量，该常量代表应用于指定填充或线条的图案。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**              | 返回图形填充格式的类型。只读 **MsoFillType** 类型。          |

**成员方法**

#### **FillFormat.OneColorGradient**

将指定填充设置为单色渐变。

**语法**

**express.OneColorGradient(Style, Variant, Degree)**

*express*   一个代表 **FillFormat** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**         | **说明**                                                     |
| --------- | ------------- | -------------------- | ------------------------------------------------------------ |
| *Style*   | 必选          | **MsoGradientStyle** | 渐变样式。可以是任意 MsoGradientStyle 常量（msoGradientFromTitle 除外，它仅应用于 WPP）。 |
| *Variant* | 必选          | **Number**           | 渐变变量。取值范围为从 1 到 4，对应于“填充效果”对话框中“渐变”选项卡上的四种变形。如果 Style 为 msoGradientFromCenter，则该参数可为 1 或 2。 Degree 必选 Single 渐变深度。值的范围从 0.0（深）到 1.0（浅）。 |
| *Degree*  | 必选          | **Number**           | 渐变深度。值的范围从 0.0（深）到 1.0（浅）。                 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将在活动文档中添加一个具有单色渐变填充效果的矩形 */ function test() { 	let shp = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80); 	let fill = shp.Fill; 	fill.ForeColor.RGB = 0x00FFFF 	fill.OneColorGradient(msoGradientFromCenter, 3, 1); } ` |

#### **FillFormat.Patterned**

将指定填充设置为一种图案

**语法**

**express.Patterned(Pattern)**

*express*   一个代表 **FillFormat** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**       | **说明**                 |
| --------- | ------------- | ------------------ | ------------------------ |
| *Pattern* | 必选          | **MsoPatternType** | 用于指定填充方式的图案。 |

**说明**

可用 **BackColor** 和 **ForeColor** 属性设置图案颜色。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在活动文档中添加一个作为填充图案的矩形。 */ function test() { 	let shp = Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80); 	let fill = shp.Fill; 	fill.Patterned(msoPatternDarkVertical); }` |

#### **FillFormat.Solid**

将指定填充设置为统一的颜色。

**语法**

**express.Solid()**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

使用该方法将渐变、纹理、图案或背景填充转换为纯色填充。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 将当前文档第一个图片设置为绿色单色填充 */ function test() { 	let firstShape = Application.ActiveDocument.Shapes.Item(1); 	firstShape.Fill.Solid() 	firstShape.Fill.ForeColor.RGB = 0x00ff00 }` |

#### **FillFormat.TwoColorGradient**

将指定填充设为双色渐变。

**语法**

**express.TwoColorGradient(Style, Variant)**

*express*   一个代表 **FillFormat** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**         | **说明**                                                     |
| --------- | ------------- | -------------------- | ------------------------------------------------------------ |
| *Style*   | 必选          | **MsoGradientStyle** | 渐变样式。                                                   |
| *Variant* | 必选          | **Number**           | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 Style 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |

#### **FillFormat.UserPicture**

用图像填充指定的形状。

**语法**

**express.UserPicture(PictureFile)**

*express*   一个代表 **FillFormat** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**     |
| ------------- | ------------- | ------------ | ------------ |
| *PictureFile* | 必选          | **String**   | 图片文件名。 |

**成员属性**

#### **FillFormat.BackColor**

返回或设置一个 **ColorFormat** 对象，该对象代表填充的背景色，可读写。

**语法**

**express.BackColor**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例添加一个矩形，并返回填充背景色 */ Application.ActiveDocument.Shapes.AddShape(msoShapeRectangle,10,10,20,30).Fill.BackColor` |

#### **FillFormat.ForeColor**

返回或设置一个 **ColorFormat** 对象，该对象代表填充的前景色。可读写。

**语法**

**express.ForeColor**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 设置矩形填充的前景色 */ Application.ActiveDocument.Shapes.Item(1).Fill.ForeColor.RGB=0x0000ff` |

#### **FillFormat.GradientAngle**

返回或设置指定填充格式的渐变填充的角度。可读/写。

**语法**

**express.GradientAngle**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

可以在图表中各种元素（形状）的格式设置中指定渐变填充。例如，可以使用**“数据系列格式”**对话框将**“柱形图”**图表中列的格式设置为渐变填充。在这种情况下，**GradientAngle** 属性与**“数据系列格式”**对话框的**“填充”**类别中**“角度”**框的设置相对应。**GradientAngle** 属性的有效值范围为从 0 到 359.9。 

#### **FillFormat.GradientColorType**

返回指定填充的渐变颜色的样式。只读 **MsoGradientColorType** 类型。

**语法**

**express.GradientColorType**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

此属性为只读属性。可用 **OneColorGradient**、**PresetGradient** 或 **TwoColorGradient** 方法设置填充的渐变类型。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将活动文档中所有图形的双色渐变填充更改为预设的渐变填充 */ function test() { 	let shps = Application.ActiveDocument.Shapes; 	for (let i = 1; i <= shps.Count; i++) { 		if (shps.Item(i).Fill.GradientColorType === msoGradientColorMixed) { 			shps.Item(i).Fill.PresetGradient(msoGradientFromCenter,1,msoGradientBrass); 		} 	} }` |

#### **FillFormat.GradientDegree**

返回值表明单色渐变填充的深浅程度。**Number** 类型，只读。

**语法**

**express.GradientDegree**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

0 表示黑色与图形前景色混色形成渐变；1 表示白色混色；介于 0 和 1 之间的值表示混入前景色偏深或偏白的阴影。

可用 **OneColorGradient** 方法设置填充的渐变程度。

#### **FillFormat.Pattern**

返回或设置一个 **MsoPatternType** 常量，该常量代表应用于指定填充或线条的图案。只读。

**语法**

**express.Pattern**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 获取激活文档的第一个填充图案 */ Application.ActiveDocument.Shapes.Item(1).Fill.Pattern` |

#### **FillFormat.Type**

返回图形填充格式的类型。只读 **MsoFillType** 类型。

**语法**

**express.Type**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 获取当前文档填充类型 */ Application.ActiveDocument.Shapes.Item(1).Fill.Type` |

适用环境：web

适用平台：windows/linux