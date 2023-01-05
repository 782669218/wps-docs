#### **ShadowFormat**



代表形状的阴影格式。

**说明**

使用 **Shadow** 属性可返回一个 **ShadowFormat** 对象。以下示例向活动文档添加一个带阴影的矩形。半透明的蓝色阴影超出矩形右边 10 磅，高出矩形 10 磅。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let shps = Application.ActiveDocument.Shapes; 	let shp = shps.AddShape(msoShapeRectangle,50,50,100,200); 	let shd = shp.Shadow; 	shd.ForeColor.RGB=0x7f0000; 	shd.OffsetX = 10; 	shd.OffsetY = 10; 	shd.Size = 100; 	shd.Transparency = 0.5; 	shd.Visible = true; }` |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Blur**            | 返回或设置一个 **Number** 类型的值，该值代表阴影格式的模糊级别。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ForeColor**       | 返回或设置一个 **ColorFormat** 对象，该对象代表填充、线条或阴影的前景色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OffsetX**         | 返回或设置阴影根据指定图形的水平偏移量（以磅为单位）。正值表示阴影向图形右侧偏移，负值表示向左偏移。**Number** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OffsetY**         | 返回或设置阴影相对指定图形的垂直偏移量（以磅为单位）。**Number** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RotateWithShape** | 返回或设置一个 **MsoTriState**，代表在旋转形象时是否旋转阴影。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Size**            | 返回或设置一个 **Number** 类型的值，该值代表阴影的宽度。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Transparency**    | 返回或设置指定阴影的透明度，其值为从 0.0（不透明）到 1.0（全透明）。可读写 **Number** 类型。 |

**成员属性**

#### **ShadowFormat.Blur**

返回或设置一个 **Number** 类型的值，该值代表阴影格式的模糊级别。可读写。

**语法**

**express.Blur**

*express*   一个代表 **ShadowFormat** 对象的变量。

#### **ShadowFormat.ForeColor**

返回或设置一个 **ColorFormat** 对象，该对象代表填充、线条或阴影的前景色。可读写。

**语法**

**express.ForeColor**

*express*   一个代表 **ShadowFormat** 对象的变量。

#### **ShadowFormat.OffsetX**

返回或设置阴影根据指定图形的水平偏移量（以磅为单位）。正值表示阴影向图形右侧偏移，负值表示向左偏移。**Number** 类型，可读写。

**语法**

**express.OffsetX**

*express*   一个代表 **ShadowFormat** 对象的变量。

**说明**

如果不指明绝对位置又要从当前位置水平或垂直地微移阴影，请使用 **IncrementOffsetX** 或 **IncrementOffsetY** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let shps = Application.ActiveDocument.Shapes; 	let shp = shps.AddShape(msoShapeRectangle,50,50,100,200); 	let shd = shp.Shadow; 	shd.ForeColor.RGB=0x7f0000; 	shd.OffsetX = 10; 	shd.OffsetY = 10; 	shd.Size = 100; 	shd.Transparency = 0.5; 	shd.Visible = true; }` |

#### **ShadowFormat.OffsetY**

返回或设置阴影相对指定图形的垂直偏移量（以磅为单位）。**Number** 类型，可读写。

**语法**

**express.OffsetY**

*express*   一个代表 **ShadowFormat** 对象的变量。

**说明**

正值代表阴影向图形下偏移，负值代表向上偏移。如果不指明绝对位置又要从当前位置水平或垂直地微移阴影，请使用 **IncrementOffsetX** 或 **IncrementOffsetY** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let shps = Application.ActiveDocument.Shapes; 	let shp = shps.AddShape(msoShapeRectangle,50,50,100,200); 	let shd = shp.Shadow; 	shd.ForeColor.RGB=0x7f0000; 	shd.OffsetX = 10; 	shd.OffsetY = 10; 	shd.Size = 100; 	shd.Transparency = 0.5; 	shd.Visible = true; }` |

#### **ShadowFormat.RotateWithShape**

返回或设置一个 **MsoTriState**，代表在旋转形象时是否旋转阴影。可读写。

**语法**

**express.RotateWithShape**

*express*   一个代表 **ShadowFormat** 对象的变量。

#### **ShadowFormat.Size**

返回或设置一个 **Number** 类型的值，该值代表阴影的宽度。可读写。

**语法**

**express.Size**

*express*   一个代表 **ShadowFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 设置阴影尺寸为10 */ Application.ActiveDocument.Shapes.Item(1).Shadow.Size = 10` |

#### **ShadowFormat.Transparency**

返回或设置指定阴影的透明度，其值为从 0.0（不透明）到 1.0（全透明）。可读写 **Number** 类型。

**语法**

**express.Transparency**

*express*   一个代表 **ShadowFormat** 对象的变量。

适用环境：web

适用平台：windows/linux