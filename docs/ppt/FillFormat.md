**FillFormat**



代表形状的填充格式。形状可以有单色、渐变、纹理、图案、图片或半透明填充。

**说明**

**FillFormat** 对象的许多属性都是只读的。若要设置其中一个属性，必须应用相应的方法。

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Background**       | 指定形状的填充与幻灯片背景相匹配。如果将此方法应用于填充后又更改了幻灯片背景色，该填充会随之改变。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OneColorGradient** | 将指定填充设置为单色渐变。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Patterned**        | 将指定填充设置为一种图案。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PresetGradient**   | 将指定填充设为一个预设的渐变。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PresetTextured**   | 将指定的填充方式设置为预设纹理。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Solid**            | 将指定的填充格式设置为均匀的颜色。可用本方法将带有渐变、纹理、图案或背景的填充格式转换为单色的填充格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TwoColorGradient** | 设置指定填充为双色渐变。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UserPicture**      | 用一幅大图像来填充指定形状。如果要通过平铺小图像来填充形状，请使用 **UserTextured** 方法。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UserTextured**     | 用平铺的小图像填充指定形状。如果要用一个大图像来填充该形状，请使用 **UserPicture** 方法。 |

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**        | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackColor**          | 返回或设置 **ColorFormat** 对象，该对象代表指定填充对象或图案线条的背景色。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**            | 返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ForeColor**          | 返回或设置 **ColorFormat** 对象，该对象表示填充、线条或阴影的前景色。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GradientColorType**  | 返回指定填充的渐变颜色类型。此属性为只读。使用 **OneColorGradient**、**PresetGradient** 或 **TwoColorGradient** 方法设置填充的渐变类型。只读 **MsoGradientColorType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GradientDegree**     | 返回一个值，该值指示单色渐变填充的深浅程度。该值为 0（零）表示形状的前景色中混入黑色形成渐变；值为 1 表示混入白色；介于 0 和 1 之间的值表示混入前景色的偏深或偏浅的底纹。只读 **Single** 类型。此属性为只读。使用 **OneColorGradient** 方法设置填充的渐变程度。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GradientStyle**      | 返回指定填充的渐变样式。使用 **OneColorGradient**、**PresetGradient** 或 **TwoColorGradient** 方法设置填充的渐变样式。尝试对未使用渐变的填充返回此属性将产生错误。使用 **Type** 属性决定填充是否使用渐变。只读 **MsoGradientStyle** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GradientVariant**    | 返回指定填充的渐变变量，对于大多数渐变填充，其返回值为从 1 到 4 的整数。如果渐变样式为 **msoGradientFromTitle** 或 **msoGradientFromCenter**，则该属性返回 1 或 2。该属性的值与**“形状填充”**选项卡中的**“渐变”**子选项卡上的渐变变量（从左向右，从上向下编号）相对应。**Long** 类型，只读。此属性为只读。使用 **OneColorGradient**、**PresetGradient** 或 **TwoColorGradient** 方法可设置填充的渐变变量。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**             | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Pattern**            | 设置或返回一个值，该值代表应用于指定填充的图案。使用 **BackColor** 和 **ForeColor** 属性可设置图案中使用的颜色。只读。**MsoPatternType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PresetGradientType** | 返回指定填充的预设渐变类型。只读 **MsoPresetGradientType** 类型。使用 **PresetGradient** 方法可设置填充的预设渐变类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PresetTexture**      | 返回指定填充的预设纹理。只读 **MsoPresetTexture** 类型。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextureName**        | 返回指定填充的自定义纹理文件名称。只读 **String** 类型。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextureType**        | 返回指定填充的纹理类型。**MsoTextureType** 类型，只读。使用 **PresetTextured** 或 **UserTextured** 方法可设置填充的纹理类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Transparency**       | 该属性返回或设置指定填充、阴影或线条的透明度，可取 0.0（不透明）至 1.0（透明）之间的数值。**Single** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**               | 返回一个 **MsoFillType** 常量，该常量代表填充的类型。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**            | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。 |

**成员方法**

#### **FillFormat.Background**

指定形状的填充与幻灯片背景相匹配。如果将此方法应用于填充后又更改了幻灯片背景色，该填充会随之改变。

**语法**

**express.Background()**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

请注意，将 **Background** 方法应用于形状的填充与设置该形状的透明填充不完全相同，也不完全等同于将形状填充设为与背景相同。请参见第二个示例。

**示例**

本示例将当前演示文稿中第一张幻灯片第一个形状的填充设为与幻灯片背景相匹配。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.Background()` |

本示例将当前演示文稿第一张幻灯片的背景设为一个预设的渐变填充，在幻灯片中添加一个矩形，然后在该矩形前面放置三个椭圆。第一个椭圆的填充与幻灯片背景相匹配，第二个的填充为透明，而第三个使用的填充方式与背景相同。请注意三个椭圆外观的不同。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ 　　　　let aSlides = Application.ActivePresentation.Slides.Item(1) 　　　　aSlides.FollowMasterBackground = false 　　　　aSlides.Background.Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientDaybreak) 　　　　let dBackground = aSlides.Shapes 　　　　dBackground.AddShape(msoShapeRectangle, 50, 200, 600, 100) 　　　　dBackground.AddShape(msoShapeOval, 75, 150, 150, 100).Fill.Background() 　　　　dBackground.AddShape(msoShapeOval, 275, 150, 150, 100).Fill.Transparency = 1 　　　　dBackground.AddShape(msoShapeOval, 475, 150, 150, 100).Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientDaybreak) }` |

#### **FillFormat.OneColorGradient**

将指定填充设置为单色渐变。

**语法**

**express.OneColorGradient(Style, Variant, Degree)**

*express*   一个代表 **FillFormat** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**         | **说明**                                                     |
| --------- | ------------- | -------------------- | ------------------------------------------------------------ |
| *Style*   | 必选          | **MsoGradientStyle** | 渐变样式。                                                   |
| *Variant* | 必选          | **Long**             | 渐变变量。取值范围为 1 到 4，对应于“形状填充”选项卡中“渐变”选项卡上的四个渐变变量。如果 Style 为 msoGradientFromTitle 或 msoGradientFromCenter，则该参数可设为 1 或 2。 |
| *Degree*  | 必选          | **Single**           | 渐变程度。可为 0.0（深）到 1.0（浅）之间的值。               |

**说明**

| MsoGradientStyle 可以是下列 MsoGradientStyle 常量之一。 |
| ------------------------------------------------------- |
| msoGradientDiagonalDown                                 |
| msoGradientDiagonalUp                                   |
| msoGradientFromCenter                                   |
| msoGradientFromCorner                                   |
| msoGradientFromTitle                                    |
| msoGradientHorizontal                                   |
| msoGradientMixed                                        |
| msoGradientVertical                                     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加一个具有单色渐变填充效果的矩形。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill 　　　　dShapes.ForeColor.RGB = (0, 128, 128) 　　　　dShapes.OneColorGradient(msoGradientHorizontal, 1, 1) }` |

#### **FillFormat.Patterned**

将指定填充设置为一种图案。

**语法**

**express.Patterned(Pattern)**

*express*   一个代表 **FillFormat** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**       | **说明**                 |
| --------- | ------------- | ------------------ | ------------------------ |
| *Pattern* | 必选          | **MsoPatternType** | 用于指定填充方式的图案。 |

**说明**

使用 **BackColor** 和 **ForeColor** 属性可设置图案中使用的颜色。

| MsoPatternType 可以是下列 MsoPatternType 常量之一。 |
| --------------------------------------------------- |
| msoPattern10Percent                                 |
| msoPattern20Percent                                 |
| msoPattern25Percent                                 |
| msoPattern30Percent                                 |
| msoPattern40Percent                                 |
| msoPattern50Percent                                 |
| msoPattern5Percent                                  |
| msoPattern60Percent                                 |
| msoPattern70Percent                                 |
| msoPattern75Percent                                 |
| msoPattern80Percent                                 |
| msoPattern90Percent                                 |
| msoPatternDarkDownwardDiagonal                      |
| msoPatternDarkHorizontal                            |
| msoPatternDarkUpwardDiagonal                        |
| msoPatternDashedDownwardDiagonal                    |
| msoPatternDashedHorizontal                          |
| msoPatternDashedUpwardDiagonal                      |
| msoPatternDashedVertical                            |
| msoPatternDiagonalBrick                             |
| msoPatternDivot                                     |
| msoPatternDottedDiamond                             |
| msoPatternDottedGrid                                |
| msoPatternHorizontalBrick                           |
| msoPatternLargeCheckerBoard                         |
| msoPatternLargeConfetti                             |
| msoPatternLargeGrid                                 |
| msoPatternLightDownwardDiagonal                     |
| msoPatternLightHorizontal                           |
| msoPatternLightUpwardDiagonal                       |
| msoPatternLightVertical                             |
| msoPatternMixed                                     |
| msoPatternNarrowHorizontal                          |
| msoPatternNarrowVertical                            |
| msoPatternOutlinedDiamond                           |
| msoPatternPlaid                                     |
| msoPatternShingle                                   |
| msoPatternSmallCheckerBoard                         |
| msoPatternSmallConfetti                             |
| msoPatternSmallGrid                                 |
| msoPatternSolidDiamond                              |
| msoPatternSphere                                    |
| msoPatternTrellis                                   |
| msoPatternWave                                      |
| msoPatternWeave                                     |
| msoPatternWideDownwardDiagonal                      |
| msoPatternWideUpwardDiagonal                        |
| msoPatternZigZag                                    |
| msoPatternDarkVertical                              |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一个带图案填充的椭圆添加到 myDocument 中。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes.AddShape(msoShapeOval, 60, 60, 80, 40).Fill 　　　　dShapes.ForeColor.RGB = (128, 0, 0) 　　　　dShapes.BackColor.RGB = (0, 0, 255) 　　　　dShapes.Patterned(msoPatternDarkVertical) }` |

#### **FillFormat.PresetGradient**

将指定填充设为一个预设的渐变。

**语法**

**express.PresetGradient(Style, Variant, PresetGradientType)**

*express*   一个代表 **FillFormat** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型**              | **说明**                                                     |
| -------------------- | ------------- | ------------------------- | ------------------------------------------------------------ |
| *Style*              | 必选          | **MsoGradientStyle**      | 渐变样式。                                                   |
| *Variant*            | 必选          | **Integer**               | 渐变变量。取值范围为 1 到 4，对应于“形状填充”选项卡的“渐变”子选项卡上的四个渐变变量。如果 Style 为 msoGradientFromTitle 或 msoGradientFromCenter，则该参数可设为 1 或 2。 |
| *PresetGradientType* | 必选          | **MsoPresetGradientType** | 渐变类型。                                                   |

**说明**

| MsoGradientStyle 可以是下列 MsoGradientStyle 常量之一。 |
| ------------------------------------------------------- |
| msoGradientDiagonalDown                                 |
| msoGradientDiagonalUp                                   |
| msoGradientFromCenter                                   |
| msoGradientFromCorner                                   |
| msoGradientFromTitle                                    |
| msoGradientHorizontal                                   |
| msoGradientMixed                                        |
| msoGradientVertical                                     |

| MsoPresetGradientType 可以是下列 MsoPresetGradientType 常量之一。 |
| ------------------------------------------------------------ |
| msoGradientBrass                                             |
| msoGradientCalmWater                                         |
| msoGradientChrome                                            |
| msoGradientChromeII                                          |
| msoGradientDaybreak                                          |
| msoGradientDesert                                            |
| msoGradientEarlySunset                                       |
| msoGradientFire                                              |
| msoGradientFog                                               |
| msoGradientGold                                              |
| msoGradientGoldII                                            |
| msoGradientHorizon                                           |
| msoGradientLateSunset                                        |
| msoGradientMahogany                                          |
| msoGradientMoss                                              |
| msoGradientNightfall                                         |
| msoGradientOcean                                             |
| msoGradientParchment                                         |
| msoGradientPeacock                                           |
| msoGradientRainbow                                           |
| msoGradientRainbowII                                         |
| msoGradientSapphire                                          |
| msoGradientSilver                                            |
| msoGradientWheat                                             |
| msoPresetGradientMixed                                       |

**示例**

本示例将一个带有预设渐变填充的矩形添加到 `myDocument`。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.ActivePresentation.Slides.Item(1)     myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 140, 80).Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientBrass) }` |

#### **FillFormat.PresetTextured**

将指定的填充方式设置为预设纹理。

**语法**

**express.PresetTextured(PresetTexture)**

*express*   一个代表 **FillFormat** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**         | **说明**   |
| --------------- | ------------- | -------------------- | ---------- |
| *PresetTexture* | 必选          | **MsoPresetTexture** | 预设纹理。 |

**说明**

| MsoPresetTexture 可以是下列 MsoPresetTexture 常量之一。 |
| ------------------------------------------------------- |
| msoPresetTextureMixed                                   |
| msoTextureBlueTissuePaper                               |
| msoTextureBouquet                                       |
| msoTextureBrownMarble                                   |
| msoTextureCanvas                                        |
| msoTextureCork                                          |
| msoTextureDenim                                         |
| msoTextureFishFossil                                    |
| msoTextureGranite                                       |
| msoTextureGreenMarble                                   |
| msoTextureMediumWood                                    |
| msoTextureNewsprint                                     |
| msoTextureOak                                           |
| msoTexturePaperBag                                      |
| msoTexturePapyrus                                       |
| msoTextureParchment                                     |
| msoTexturePinkTissuePaper                               |
| msoTexturePurpleMesh                                    |
| msoTextureRecycledPaper                                 |
| msoTextureSand                                          |
| msoTextureStationery                                    |
| msoTextureWalnut                                        |
| msoTextureWaterDroplets                                 |
| msoTextureWhiteMarble                                   |
| msoTextureWovenMat                                      |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一个带有灰褐色大理石纹理填充的圆柱形添加到 myDocument。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　myDocument.Shapes.AddShape(msoShapeCan, 90, 90, 40, 80).Fill.PresetTextured(msoTextureGreenMarble) }` |

#### **FillFormat.Solid**

将指定的填充格式设置为均匀的颜色。可用本方法将带有渐变、纹理、图案或背景的填充格式转换为单色的填充格式。

**语法**

**express.Solid()**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 上的所有填充转换为统一的红色填充。*/ function test() {     let myDocument = Application.ActivePresentation.Slides.Item(1)     for (let s = 1; s <= myDocument.Shapes.Count; s++) {         myDocument.Shapes.Item(s).Fill.Solid()         myDocument.Shapes.Item(s).Fill.ForeColor.RGB = (0, 255, 255)     } }　　 ` |

#### **FillFormat.TwoColorGradient**

设置指定填充为双色渐变。

**语法**

**express.TwoColorGradient(Style, Variant)**

*express*   一个代表 **FillFormat** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**         | **说明**                                                     |
| --------- | ------------- | -------------------- | ------------------------------------------------------------ |
| *Style*   | 必选          | **MsoGradientStyle** | 渐变样式。                                                   |
| *Variant* | 必选          | **Long**             | 渐变变量。取值范围为 1 到 4，对应于“形状填充”选项卡的“渐变”子选项卡上的四个渐变变量。如果 Style 为 msoGradientFromTitle 或 msoGradientFromCenter，则该参数可设为 1 或 2。 |

**说明**

| MsoGradientStyle 可以是下列 MsoGradientStyle 常量之一。 |
| ------------------------------------------------------- |
| msoGradientDiagonalDown                                 |
| msoGradientDiagonalUp                                   |
| msoGradientFromCenter                                   |
| msoGradientFromCorner                                   |
| msoGradientFromTitle                                    |
| msoGradientHorizontal                                   |
| msoGradientMixed                                        |
| msoGradientVertical                                     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加一个具有双色渐变填充效果的矩形，并设置填充的前景色和背景色。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill 　　　　dShapes.ForeColor.RGB = (128, 0, 0) 　　　　dShapes.BackColor.RGB = (0, 170, 170) 　　　　dShapes.TwoColorGradient(msoGradientHorizontal, 1) }` |

#### **FillFormat.UserPicture**

用一幅大图像来填充指定形状。如果要通过平铺小图像来填充形状，请使用 **UserTextured** 方法。

**语法**

**express.UserPicture(PictureFile)**

*express*   一个代表 **FillFormat** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**     |
| ------------- | ------------- | ------------ | ------------ |
| *PictureFile* | 必选          | **String**   | 图片文件名。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加了两个矩形。用 Tiles.bmp 文件中的一幅大图像填充左边的矩形；用 Tiles.bmp 文件中若干平铺的小图像填充右边的矩形*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes 　　　　dShapes.AddShape(msoShapeRectangle, 0, 0, 200, 100).Fill.UserPicture("c:\\windows\\tiles.bmp") 　　　　dShapes.AddShape(msoShapeRectangle, 300, 0, 200, 100).Fill.UserTextured("c:\\windows\\tiles.bmp") }` |

#### **FillFormat.UserTextured**

用平铺的小图像填充指定形状。如果要用一个大图像来填充该形状，请使用 **UserPicture** 方法。

**语法**

**express.UserTextured()**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加了两个矩形。用 Tiles.bmp 文件中的一幅大图像填充左边的矩形；用 Tiles.bmp 文件中若干平铺的小图像填充右边的矩形*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes 　　　　dShapes.AddShape(msoShapeRectangle, 0, 0, 200, 100).Fill.UserPicture("c:\\windows\\tiles.bmp") 　　　　dShapes.AddShape(msoShapeRectangle, 300, 0, 200, 100).Fill.UserTextured("c:\\windows\\tiles.bmp") }` |

**成员属性**

#### **FillFormat.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `funciton AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **FillFormat.BackColor**

返回或设置 **ColorFormat** 对象，该对象代表指定填充对象或图案线条的背景色。可读/写。

**语法**

**express.BackColor**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变。*/ function test(){ 　　　　let myDocument = ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill 　　　　dShapes.ForeColor.RGB = (128, 0, 0) 　　　　dShapes.BackColor.RGB = (170, 170, 170) 　　　　dShapes.TwoColorGradient(msoGradientHorizontal, 1) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将一条带图案的线条添加至 myDocument。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dLine = myDocument.Shapes.AddLine(10, 100, 250, 0).Line 　　　　dLine.Weight = 6 　　　　dLine.ForeColor.RGB = (0, 0, 255) 　　　　dLine.BackColor.RGB = (128, 0, 0) 　　　　dLine.Pattern = msoPatternDarkDownwardDiagonal }` |

#### **FillFormat.Creator**

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示关于 myObject 创建者的消息。*/ function test(){ 　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1) 　　　　if ( myObject.Creator == 0x50575054 ) { 　　　　    alert("This is a WPP object") 　　　　} 　　　　else { 　　　　    alert("This is not a WPP object") 　　　　} }` |

#### **FillFormat.ForeColor**

返回或设置 **ColorFormat** 对象，该对象表示填充、线条或阴影的前景色。可读/写。

**语法**

**express.ForeColor**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加一个矩形，然后设置该矩形的填充的前景色、背景色和渐变。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill 　　　　dShapes.ForeColor.RGB = (128, 0, 0) 　　　　dShapes.BackColor.RGB = (170, 170, 170) 　　　　dShapes.TwoColorGradient(msoGradientHorizontal, 1)}` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将一条带图案的线条添加至 myDocument。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dLine = myDocument.Shapes.AddLine(10, 100, 250, 0).Line 　　　　dLine.Weight = 6 　　　　dLine.ForeColor.RGB = (0, 0, 255) 　　　　dLine.BackColor.RGB = (128, 0, 0) 　　　　dLine.Pattern = msoPatternDarkDownwardDiagonal }` |

#### **FillFormat.GradientColorType**

返回指定填充的渐变颜色类型。此属性为只读。使用 **OneColorGradient**、**PresetGradient** 或 **TwoColorGradient** 方法设置填充的渐变类型。只读 **MsoGradientColorType** 类型。

**语法**

**express.GradientColorType**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

| MsoGradientColorType 可为以下 MsoGradientColorType 常量之一。 |
| ------------------------------------------------------------ |
| msoGradientColorMixed                                        |
| msoGradientOneColor                                          |
| msoGradientPresetColors                                      |
| msoGradientTwoColors                                         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中使用双色渐变填充的所有形状的填充更改为预设的渐变填充。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　for ( let s = 1; s <= myDocument.Shapes.Count; s++ ) { 　　　　    if ( myDocument.Shapes.Item(s).Fill.GradientColorType == msoGradientTwoColors ) { 　　　　        myDocument.Shapes.Item(s).Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientBrass) 　　　　    } 　　　　}} ` |

#### **FillFormat.GradientDegree**

返回一个值，该值指示单色渐变填充的深浅程度。该值为 0（零）表示形状的前景色中混入黑色形成渐变；值为 1 表示混入白色；介于 0 和 1 之间的值表示混入前景色的偏深或偏浅的底纹。只读 **Single** 类型。此属性为只读。使用 **OneColorGradient** 方法设置填充的渐变程度。

**语法**

**express.GradientDegree**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 添加矩形，并设置该矩形的填充渐变程度与名为“Rectangle 2”的形状中的渐变程度一致。如果“Rectangle 2”未使用单色渐变填充，则本示例失败。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes 　　　　let gradDegree1 = dShapes.Item("Rectangle 2").Fill.GradientDegree 　　　　let filldShapes = dShapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill 　　　　filldShapes.ForeColor.RGB = (128, 0, 0) 　　　　filldShapes.OneColorGradient(msoGradientHorizontal, 1, gradDegree1) }` |

#### **FillFormat.GradientStyle**

返回指定填充的渐变样式。使用 **OneColorGradient**、**PresetGradient** 或 **TwoColorGradient** 方法设置填充的渐变样式。尝试对未使用渐变的填充返回此属性将产生错误。使用 **Type** 属性决定填充是否使用渐变。只读 **MsoGradientStyle** 类型。

**语法**

**express.GradientStyle**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

| MsoGradientStyle 可以是下列 MsoGradientStyle 常量之一。 |
| ------------------------------------------------------- |
| msoGradientDiagonalDown                                 |
| msoGradientDiagonalUp                                   |
| msoGradientFromCenter                                   |
| msoGradientFromCorner                                   |
| msoGradientFromTitle                                    |
| msoGradientHorizontal                                   |
| msoGradientMixed                                        |
| msoGradientVertical                                     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加矩形，并设置该矩形的填充渐变样式与名为“rect1”的形状的渐变样式一致。“rect1”必须具有渐变填充，本示例才有效。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes 　　　　let gradStyle1 = dShapes.Item("rect1").Fill.GradientStyle 　　　　let filldShapes = dShapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill 　　　　filldShapes.ForeColor.RGB = (128, 0, 0) 　　　　filldShapes.OneColorGradient(gradStyle1, 1, 1) }` |

#### **FillFormat.GradientVariant**

返回指定填充的渐变变量，对于大多数渐变填充，其返回值为从 1 到 4 的整数。如果渐变样式为 **msoGradientFromTitle** 或 **msoGradientFromCenter**，则该属性返回 1 或 2。该属性的值与**“形状填充”**选项卡中的**“渐变”**子选项卡上的渐变变量（从左向右，从上向下编号）相对应。**Long** 类型，只读。此属性为只读。使用 **OneColorGradient**、**PresetGradient** 或 **TwoColorGradient** 方法可设置填充的渐变变量。

**语法**

**express.GradientVariant**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向 myDocument 中添加矩形，并设置该矩形的填充渐变变量，以匹配名为“rect1”的形状的渐变变量。“rect1”必须具有渐变填充，本示例才有效。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes 　　　　let gradVar1 = dShapes.Item("rect1").Fill.GradientVariant 　　　　let filldShapes = dShapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill 　　　　filldShapes.ForeColor.RGB = (128, 0, 0) }` |

#### **FillFormat.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **FillFormat.Pattern**

设置或返回一个值，该值代表应用于指定填充的图案。使用 **BackColor** 和 **ForeColor** 属性可设置图案中使用的颜色。只读。**MsoPatternType** 类型。

**语法**

**express.Pattern**

*express*   一个代表 **FillFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加一个矩形，并设置其填充图案与名为“rect1”的形状保持一致。新矩形使用和 rect1 相同的图案，但颜色未必相同。通过 BackColor 和 ForeColor 属性可设置该图案中使用的颜色。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes 　　　　let pattern1 = dShapes.Item("rect1").Fill.Pattern 　　　　let fillShapes = dShapes.AddShape(msoShapeRectangle, 100, 100, 120, 80).Fill 　　　　fillShapes.ForeColor.RGB = (128, 0, 0) 　　　　fillShapes.BackColor.RGB = (0, 0, 255) 　　　　fillShapes.Patterned(pattern1) }` |

#### **FillFormat.PresetGradientType**

返回指定填充的预设渐变类型。只读 **MsoPresetGradientType** 类型。使用 **PresetGradient** 方法可设置填充的预设渐变类型。

**语法**

**express.PresetGradientType**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

| MsoPresetGradientType 可以是下列 MsoPresetGradientType 常量之一。 |
| ------------------------------------------------------------ |
| msoGradientBrass                                             |
| msoGradientCalmWater                                         |
| msoGradientChrome                                            |
| msoGradientChromeII                                          |
| msoGradientDaybreak                                          |
| msoGradientDesert                                            |
| msoGradientEarlySunset                                       |
| msoGradientFire                                              |
| msoGradientFog                                               |
| msoGradientGold                                              |
| msoGradientGoldII                                            |
| msoGradientHorizon                                           |
| msoGradientLateSunset                                        |
| msoGradientMahogany                                          |
| msoGradientMoss                                              |
| msoGradientNightfall                                         |
| msoGradientOcean                                             |
| msoGradientParchment                                         |
| msoGradientPeacock                                           |
| msoGradientRainbow                                           |
| msoGradientRainbowII                                         |
| msoGradientSapphire                                          |
| msoGradientSilver                                            |
| msoGradientWheat                                             |
| msoPresetGradientMixed                                       |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 myDocument 中所有具有 Moss 预设渐变填充的形状的填充更改为 Fog 预设渐变填充。*/ function test(){  　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　for ( let s = 1; s <= myDocument.Shapes.Count; s++ ) { 　　　　    if ( myDocument.Shapes.Item(s).Fill.PresetGradientType == msoGradientMoss ) { 　　　　        myDocument.Shapes.Item(s).Fill.PresetGradient = msoGradientFog 　　　　    } 　　　　} }` |

#### **FillFormat.PresetTexture**

返回指定填充的预设纹理。只读 **MsoPresetTexture** 类型。

**语法**

**express.PresetTexture**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

| MsoPresetTexture 可以是下列 MsoPresetTexture 常量之一。 |
| ------------------------------------------------------- |
| msoPresetTextureMixed                                   |
| msoTextureBlueTissuePaper                               |
| msoTextureBouquet                                       |
| msoTextureBrownMarble                                   |
| msoTextureCanvas                                        |
| msoTextureCork                                          |
| msoTextureDenim                                         |
| msoTextureFishFossil                                    |
| msoTextureGranite                                       |
| msoTextureGreenMarble                                   |
| msoTextureMediumWood                                    |
| msoTextureNewsprint                                     |
| msoTextureOak                                           |
| msoTexturePaperBag                                      |
| msoTexturePapyrus                                       |
| msoTextureParchment                                     |
| msoTexturePinkTissuePaper                               |
| msoTexturePurpleMesh                                    |
| msoTextureRecycledPaper                                 |
| msoTextureSand                                          |
| msoTextureStationery                                    |
| msoTextureWalnut                                        |
| msoTextureWaterDroplets                                 |
| msoTextureWhiteMarble                                   |
| msoTextureWovenMat                                      |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加一个矩形，并设置该矩形的预设纹理与第二个形状的预设纹理一致。要使本示例正常运行，第二个形状必须使用预设纹理的填充。*/ function test(){ 　　　　let myDocument = Application.ActivePresentation.Slides.Item(1) 　　　　let dShapes = myDocument.Shapes 　　　　let presetTexture2 = dShapes.Item(2).Fill.PresetTexture 　　　　dShapes.AddShape(msoShapeRectangle, 100, 0, 40, 80).Fill.PresetTextured(presetTexture2)}` |

#### **FillFormat.TextureName**

返回指定填充的自定义纹理文件名称。只读 **String** 类型。

**语法**

**express.TextureName**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

此属性为只读。使用 **UserTextured** 方法可设置填充的纹理文件。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例给 myDocument 添加一个椭圆。如果 myDocument 的第一个形状使用用户定义的纹理填充，则新椭圆使用和第一个形状相同的填充。如果第一个形状使用其他任何类型的填充，则新椭圆使用绿色大理石填充*/ function test() {     let myDocument = Application.ActivePresentation.Slides.Item(1)     let newFill = myDocument.Shapes.AddShape(msoShapeOval, 0, 0, 200, 90).Fill     let fill = myDocument.Shapes.Item(1).Fill     if (fill.Type == msoFillTextured && fill.TextureType == msoTextureUserDefined) {         newFill.UserTextured(fill.TextureName)     }     else {         newFill.PresetTextured(msoTextureGreenMarble)     } } ` |

#### **FillFormat.TextureType**

返回指定填充的纹理类型。**MsoTextureType** 类型，只读。使用 **PresetTextured** 或 **UserTextured** 方法可设置填充的纹理类型。

**语法**

**express.TextureType**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

| MsoTextureType 可为以下 MsoTextureType 常量之一。 |
| ------------------------------------------------- |
| msoTexturePreset                                  |
| msoTextureTypeMixed                               |
| msoTextureUserDefined                             |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例会将 myDocument 中所有形状的填充更改为画布，这些形状全部使用自定义的纹理填充。 */ function test() {     let shape = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shape.Count; i++) {         if (shape.Item(i).Fill.TextureType == msoTextureUserDefined) {             shape.Item(i).Fill.PresetTextured(msoTextureCanvas)         }     } }` |

#### **FillFormat.Transparency**

该属性返回或设置指定填充、阴影或线条的透明度，可取 0.0（不透明）至 1.0（透明）之间的数值。**Single** 类型，可读/写。

**语法**

**express.Transparency**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本例将 myDocument 中第三个形状的阴影设置为半透明的红色。如果形状原来无阴影，以下示例为它添加阴影。*/ function test() {     let sh = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).Shadow     sh.Visible = true     sh.ForeColor.RGB = (255, 0, 0)     sh.Transparency = 0.5 }` |

#### **FillFormat.Type**

返回一个 **MsoFillType** 常量，该常量代表填充的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

| MsoFillType 可为以下 MsoFillType 常量之一。 |
| ------------------------------------------- |
| msoFillBackground                           |
| msoFillGradient                             |
| msoFillMixed                                |
| msoFillPatterned                            |
| msoFillPicture                              |
| msoFillSolid                                |
| msoFillTextured                             |

#### **FillFormat.Visible**

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。

**语法**

**express.Visible**

*express*   一个代表 **FillFormat** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| msoCTrue                                      |
| msoFalse                                      |
| msoTriStateMixed                              |
| msoTriStateToggle                             |
| msoTrue：指定对象或对象格式是可见的。         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为当前演示文稿中第一张幻灯片的第三个形状设置阴影的水平偏移量和垂直偏移量，使其向右偏移 5 磅并向上偏移 3 磅。如果该形状原来没有阴影，本例将会为其添上阴影。*/ function test() {     let sh = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).Shadow     sh.Visible = msoTrue     sh.OffsetX = 5     sh.OffsetY = -3 } ` |

适用环境：web

适用平台：windows/linux