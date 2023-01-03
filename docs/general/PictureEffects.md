#### **PictureEffects**



代表 **PictureEffects** 对象的集合。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码将设置 Wpp 幻灯片中某个形状的几个图片效果填充属性。 function test() {     // Setup a slide with one picture shape.     let pes = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects     // Insert a 150% Saturation effect.     pes.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5     // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.     let brightnessContrast = pes.Insert(Application.Enum.msoEffectBrightnessContrast)     brightnessContrast.EffectParameters.Item(1).Value = -0.5     brightnessContrast.EffectParameters.Item(2).Value = 0.25 }` |

**方法**

|                                                              | 名称       | 说明                                  |
| ------------------------------------------------------------ | ---------- | ------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 从集合中删除 **PictureEffect** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Insert** | 将图片效果插入到复合效果链中。        |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，该对象代表 **PictureEffects** 对象的容器应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 检索包含在 **PictureEffects** 集合中的 **PictureEffect** 对象数的计数。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，该整数指示在其中创建了 **PictureEffects** 对象的应用程序。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 检索位于指定索引处的 **PictureEffect** 对象。只读            |

**成员方法**

#### **PictureEffects.Delete**

从集合中删除 **PictureEffect** 对象。

**语法**

**express.Delete(Index)**

*express*   一个代表 **PictureEffects** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                              |
| -------- | ------------- | ------------ | ------------------------------------- |
| *Index*  | 可选          | **Integer**  | 要删除的 PictureEffect 对象的索引号。 |

#### **PictureEffects.Insert**

将图片效果插入到复合效果链中。

**语法**

**express.Insert(EffectType, Position)**

*express*   一个代表 **PictureEffects** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**             | **说明**                         |
| ------------ | ------------- | ------------------------ | -------------------------------- |
| *EffectType* | 必选          | **MsoPictureEffectType** | 一个指定图片效果类型的枚举。     |
| *Position*   | 可选          | **Integer**              | 效果在图片效果的复合链中的位置。 |

**说明**

图片效果被处理为由各个项构成的链，以便创建一个最终复合图像，链中的项是按照顺序应用的。效果链将允许向链中添加效果、对效果重新排序或从链中删除效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码将设置 Wpp 幻灯片中某个形状的几个图片效果填充属性。 function test() {     // Setup a slide with one picture shape.     let pes = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Fill.PictureEffects     // Insert a 150% Saturation effect.     pes.Insert(Application.Enum.msoEffectSaturation).EffectParameters.Item(1).Value = 1.5     // Insert Brightness/Contrast effect and set values to -50% Brightness and +25% Contrast.     let brightnessContrast = pes.Insert(Application.Enum.msoEffectBrightnessContrast)     brightnessContrast.EffectParameters.Item(1).Value = -0.5     brightnessContrast.EffectParameters.Item(2).Value = 0.25 }` |

**成员属性**

#### **PictureEffects.Application**

获取一个 **Application** 对象，该对象代表 **PictureEffects** 对象的容器应用程序。只读

**语法**

**express.Application**

*express*   一个代表 **PictureEffects** 对象的变量。

#### **PictureEffects.Count**

检索包含在 **PictureEffects** 集合中的 **PictureEffect** 对象数的计数。只读

**语法**

**express.Count**

*express*   一个代表 **PictureEffects** 对象的变量。

#### **PictureEffects.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **PictureEffects** 对象的应用程序。只读

**语法**

**express.Creator**

*express*   一个代表 **PictureEffects** 对象的变量。

#### **PictureEffects.Item**

检索位于指定索引处的 **PictureEffect** 对象。只读

**语法**

**express.Item**

*express*   一个代表 **PictureEffects** 对象的变量。

适用环境：web

适用平台：windows/linux