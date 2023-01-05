**ShadowFormat**



代表形状的阴影格式。

**说明**

使用 **Shadow** 属性可返回一个 **ShadowFormat** 对象。

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementOffsetX** | 按指定磅值更改阴影的水平偏移量。可以使用 **OffsetX** 属性设置阴影的绝对水平偏移量。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IncrementOffsetY** | 按指定磅值更改阴影的垂直偏移量。可以使用 **OffsetY** 属性设置阴影的绝对垂直偏移量。 |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Blur**            | 返回或设置指定底纹的模糊度。可读/写 **Single** 类型。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ForeColor**       | 返回或设置一个 **ColorFormat** 对象，它代表指定的前景填充色或纯色。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Obscured**        | 如果指定形状的阴影是填充的，并且阴影被形状所遮盖（即便该形状没有填充），那么该值为 **True**。如果阴影没有填充，并且当形状没有填充时，可透过形状看到阴影的轮廓，那么该值为 **False**。**MsoTriState** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OffsetX**         | 以磅为单位返回或设置指定形状的阴影的水平偏移量。正偏移值将阴影向右偏移，负偏移值将阴影向左偏移。可读写。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OffsetY**         | 以磅为单位返回或设置指定形状阴影的垂直偏移。正偏移值将阴影向下偏移，负偏移值将阴影向上偏移。可读写。**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**          | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RotateWithShape** | 返回或设置一个 **MsoTriState** 类型的值，该值表示是否在旋转形状时旋转阴影。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Size**            | 返回或设置指定底纹的大小。可读/写 **Single** 类型。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Style**           | 返回或设置指定底纹的样式。可读/写 **MsoShadowStyle** 类型。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Transparency**    | 返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。**Double** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**            | 返回或设置一个代表阴影格式类型的 **MsoShadowType** 值。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**         | 返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。 |

**成员方法**

#### **ShadowFormat.IncrementOffsetX**

按指定磅值更改阴影的水平偏移量。可以使用 **OffsetX** 属性设置阴影的绝对水平偏移量。

**语法**

**express.IncrementOffsetX(Increment)**

*express*   一个代表 **ShadowFormat** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定阴影的水平偏移量（以磅为单位）。正值表示阴影右移，负值表示阴影左移。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例使 myDocument 中形状三的阴影向左移动 3 磅。 function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.Item(3).Shadow.IncrementOffsetX(-3) }` |

#### **ShadowFormat.IncrementOffsetY**

按指定磅值更改阴影的垂直偏移量。可以使用 **OffsetY** 属性设置阴影的绝对垂直偏移量。

**语法**

**express.IncrementOffsetY(Increment)**

*express*   一个代表 **ShadowFormat** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Increment* | 必选          | **Single**   | 指定阴影的垂直偏移量（以磅为单位）。正值表示阴影下移，负值表示阴影上移。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例使 myDocument 中形状三的阴影向上移动 3 磅。 function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.Item(3).Shadow.IncrementOffsetY(-3) }` |

**成员属性**

#### **ShadowFormat.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ShadowFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示一条有关创建 myObject 的应用程序的消息。 function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Value == "ET")      {     	alert("This is an ET Application object.")     }     else      {     	alert("This is not an ET Application object.")     } }` |

#### **ShadowFormat.Blur**

返回或设置指定底纹的模糊度。可读/写 **Single** 类型。

**语法**

**express.Blur**

*express*   一个代表 **ShadowFormat** 对象的变量。

#### **ShadowFormat.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ShadowFormat** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ShadowFormat.ForeColor**

返回或设置一个 **ColorFormat** 对象，它代表指定的前景填充色或纯色。

**语法**

**express.ForeColor**

*express*   一个代表 **ShadowFormat** 对象的变量。

#### **ShadowFormat.Obscured**

如果指定形状的阴影是填充的，并且阴影被形状所遮盖（即便该形状没有填充），那么该值为 **True**。如果阴影没有填充，并且当形状没有填充时，可透过形状看到阴影的轮廓，那么该值为 **False**。**MsoTriState** 类型，可读写。

**语法**

**express.Obscured**

*express*   一个代表 **ShadowFormat** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。        |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse** 阴影没有填充；如果形状也没有填充，透过形状可看见阴影的轮廓。 |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 指定形状的阴影有填充并被该形状所遮蔽，即使形状也没有填充。 |

**示例**

#### **ShadowFormat.OffsetX**

以磅为单位返回或设置指定形状的阴影的水平偏移量。正偏移值将阴影向右偏移，负偏移值将阴影向左偏移。可读写。**Single** 类型。

**语法**

**express.OffsetX**

*express*   一个代表 **ShadowFormat** 对象的变量。

**说明**

使用 **IncrementOffsetX** 方法或 **IncrementOffsetY** 方法可将阴影从当前位置水平或垂直轻微移动而无需指定其绝对位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例为 myDocument 中的形状 3 设置阴影的水平和垂直偏移量。阴影相对形状向右偏移 5 磅，向上偏移 3 磅。如果形状没有阴影，本示例为其添加一个阴影。 function test() {     let shadow = Application.Worksheets.Item(1).Shapes.Item(3).Shadow     shadow.Visible = true     shadow.OffsetX = 5     shadow.OffsetY = -3 }` |

#### **ShadowFormat.OffsetY**

以磅为单位返回或设置指定形状阴影的垂直偏移。正偏移值将阴影向下偏移，负偏移值将阴影向上偏移。可读写。**Single** 类型。

**语法**

**express.OffsetY**

*express*   一个代表 **ShadowFormat** 对象的变量。

**说明**

使用 **IncrementOffsetX** 方法或 **IncrementOffsetY** 方法可将阴影从当前位置水平或垂直轻微移动而无需指定其绝对位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例为 myDocument 中的形状 3 设置阴影的水平和垂直偏移量。阴影相对形状向右偏移 5 磅，向上偏移 3 磅。如果形状没有阴影，本示例为其添加一个阴影。 function test() {     let shadow = Application.Worksheets.Item(1).Shapes.Item(3).Shadow     shadow.Visible = true     shadow.OffsetX = 5     shadow.OffsetY = -3 }` |

#### **ShadowFormat.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ShadowFormat** 对象的变量。

#### **ShadowFormat.RotateWithShape**

返回或设置一个 **MsoTriState** 类型的值，该值表示是否在旋转形状时旋转阴影。可读/写。

**语法**

**express.RotateWithShape**

*express*   一个代表 **ShadowFormat** 对象的变量。

#### **ShadowFormat.Size**

返回或设置指定底纹的大小。可读/写 **Single** 类型。

**语法**

**express.Size**

*express*   一个代表 **ShadowFormat** 对象的变量。

#### **ShadowFormat.Style**

返回或设置指定底纹的样式。可读/写 **MsoShadowStyle** 类型。

**语法**

**express.Style**

*express*   一个代表 **ShadowFormat** 对象的变量。

**说明**

| MsoShadowStyle 可以是下列 **MsoShadowStyle** 常量之一。 |
| ------------------------------------------------------- |
| **msoShadowStyleInnerShadow**                           |
| **msoShadowStyleMixed**                                 |
| **msoShadowStyleOuterShadow**                           |

#### **ShadowFormat.Transparency**

返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。**Double** 型，可读写。

**语法**

**express.Transparency**

*express*   一个代表 **ShadowFormat** 对象的变量。

**说明**

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例将第一张工作表上的第三个形状的阴影效果设置为半透明的红色。如果形状没有阴影，此示例将加上阴影效果。 function test() {     let shadow = Application.Worksheets.Item(1).Shapes.Item(3).Shadow     shadow.Visible = true     shadow.ForeColor.RGB = (255, 0, 0)     shadow.Transparency = 0.5 }` |

#### **ShadowFormat.Type**

返回或设置一个代表阴影格式类型的 **MsoShadowType** 值。

**语法**

**express.Type**

*express*   一个代表 **ShadowFormat** 对象的变量。

#### **ShadowFormat.Visible**

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **ShadowFormat** 对象的变量。

适用环境：web

适用平台：windows/linux