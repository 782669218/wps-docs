#### **EffectParameter**



描述单个图片效果参数。

**说明**

图片效果被处理为由各个项构成的链，以便创建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。效果参数指定这些效果的属性。

**示例**
下面的代码将设置 WPP 幻灯片中某个形状的几个图片效果填充属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     // Setup a slide with one picture shape.     let myEffects = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects      // Insert a 150% Saturation effect.     myEffects.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5      // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.     let brightnessContrast = myEffects.Insert(Application.Enum.msoEffectBrightnessContrast)      brightnessContrast.EffectParameters.Item(1).Value = -0.5     brightnessContrast.EffectParameters.Item(2).Value = 0.25      // Remove all Picture effects.     while(myEffects.Count > 0){         myEffects.Delete(1)     } }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，该对象代表 **EffectParameter** 对象的容器应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，该整数指示在其中创建了 **EffectParameter** 对象的应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 检索 **EffectParameter** 参数的字符串名称。只读              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**       | 检索或设置 **EffectParameter** 对象的值。可读/写             |

**成员属性**

#### **EffectParameter.Application**

获取一个 **Application** 对象，该对象代表 **EffectParameter** 对象的容器应用程序。只读

**语法**

**express.Application**

*express*   一个代表 **EffectParameter** 对象的变量。

#### **EffectParameter.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **EffectParameter** 对象的应用程序。只读

**语法**

**express.Creator**

*express*   一个代表 **EffectParameter** 对象的变量。

#### **EffectParameter.Name**

检索 **EffectParameter** 参数的字符串名称。只读

**语法**

**express.Name**

*express*   一个代表 **EffectParameter** 对象的变量。

#### **EffectParameter.Value**

检索或设置 **EffectParameter** 对象的值。可读/写

**语法**

**express.Value**

*express*   一个代表 **EffectParameter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码将 PictureEffect 对象的第一个参数设置为色温。 picEffect.EffectParameters.Item(1).Value = Application.Enum.msoEffectColorTemperature` |

适用环境：web

适用平台：windows/linux