#### **PictureEffect**



代表图片效果。

**说明**

图片效果被处理为由各个项构成的链，以便创下面的代码将设置 Wpp 幻灯片中某个形状的几个图片效果填充属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     // Setup a slide with one picture shape.     let pes = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects     // Insert a 150% Saturation effect.     pes.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5     // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.     let brightnessContrast = pes.Insert(Application.Enum.msoEffectBrightnessContrast)     brightnessContrast.EffectParameters.Item(1).Value = -0.5     brightnessContrast.EffectParameters.Item(2).Value = 0.25 }` |

建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。

**方法**

|                                                              | 名称       | 说明           |
| ------------------------------------------------------------ | ---------- | -------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除图片效果。 |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 获取一个 **Application** 对象，该对象代表 **PictureEffect** 对象的容器应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**          | 获取一个 32 位整数，该整数指示在其中创建了 **PictureEffect** 对象的应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EffectParameters** | 返回一个 **EffectParameter** 对象。只读                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**         | 指定图片效果在复合效果链中的位置。可读/写                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**             | 指定图片效果的类型。只读                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**          | 获取或设置一个代表图片效果的可视状态的 Boolean 值。可读/写   |

**成员方法**

#### **PictureEffect.Delete**

删除图片效果。

**语法**

**express.Delete()**

*express*   一个代表 **PictureEffect** 对象的变量。

**成员属性**

#### **PictureEffect.Application**

获取一个 **Application** 对象，该对象代表 **PictureEffect** 对象的容器应用程序。只读

**语法**

**express.Application**

*express*   一个代表 **PictureEffect** 对象的变量。

**说明**

#### **PictureEffect.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **PictureEffect** 对象的应用程序。只读

**语法**

**express.Creator**

*express*   一个代表 **PictureEffect** 对象的变量。

#### **PictureEffect.EffectParameters**

返回一个 **EffectParameter** 对象。只读

**语法**

**express.EffectParameters**

*express*   一个代表 **PictureEffect** 对象的变量。

**说明**

图片效果被处理为由各个项构成的链，以便创建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。效果参数指定这些效果的属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码将设置 Microsoft PowerPoint 幻灯片中某个形状的几个图片效果填充属性。 function test() {     // Setup a slide with one picture shape.     let pes = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects     // Insert a 150% Saturation effect.     pes.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5     // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.     let brightnessContrast = pes.Insert(Application.Enum.msoEffectBrightnessContrast)     brightnessContrast.EffectParameters.Item(1).Value = -0.5     brightnessContrast.EffectParameters.Item(2).Value = 0.25 }` |

#### **PictureEffect.Position**

指定图片效果在复合效果链中的位置。可读/写

**语法**

**express.Position**

*express*   一个代表 **PictureEffect** 对象的变量。

**说明**

图片效果被处理为由各个项构成的链，以便创建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码将设置 Microsoft PowerPoint 幻灯片中某个形状的几个图片效果填充属性。 function test() {     // Setup a slide with one picture shape.     let pes = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects     // Insert a 150% Saturation effect.     pes.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5     // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.     let brightnessContrast = pes.Insert(Application.Enum.msoEffectBrightnessContrast)     brightnessContrast.EffectParameters.Item(1).Value = -0.5    brightnessContrast.EffectParameters.Item(2).Value = 0.25 }` |

#### **PictureEffect.Type**

指定图片效果的类型。只读

**语法**

**express.Type**

*express*   一个代表 **PictureEffect** 对象的变量。

**说明**

此属性使用 **MsoPictureEffectType** 枚举。

#### **PictureEffect.Visible**

获取或设置一个代表图片效果的可视状态的 Boolean 值。可读/写

**语法**

**express.Visible**

*express*   一个代表 **PictureEffect** 对象的变量。

适用环境：web

适用平台：windows/linux