#### **SoftEdgeFormat**



代表形状或形状区域的软边缘格式。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS Office 应用程序的 Application对象。只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回表示用于创建指定对象的应用程序的 32 位整数。只读 **。**  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回代表指定 **SoftEdgeFormat** 对象的父对象。只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Radius**      | 返回或设置 **Single** 类型值，该值代表柔化边缘效果的辐射长度。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**        | 返回或设置一个 **MsoBevelType** 常量，该常量代表使用软边缘格式的图像边缘的棱台类型。可读写。 |

**成员属性**

#### **SoftEdgeFormat.Application**

返回一个代表 WPS Office 应用程序的 Application对象。只读。

**语法**

**express.Application**

*express*   一个代表 **SoftEdgeFormat** 对象的变量。

#### **SoftEdgeFormat.Creator**

返回表示用于创建指定对象的应用程序的 32 位整数。只读 **。**

**语法**

**express.Creator**

*express*   一个代表 **SoftEdgeFormat** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表 **string**“WPS”。该属性主要设计用于 Apple Macintosh 平台，在该平台上，每个应用程序都有一个由四个字符组成的创建者代码。例如， WPS 的创建者代码是 WPS。有关该属性的详细信息，请参考 WPS OfficeMacintosh Edition 附带的语言参考帮助。

|                                         |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **SoftEdgeFormat.Parent**

返回代表指定 **SoftEdgeFormat** 对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **SoftEdgeFormat** 对象的变量。

#### **SoftEdgeFormat.Radius**

返回或设置 **Single** 类型值，该值代表柔化边缘效果的辐射长度。可读/写。

**语法**

**express.Radius**

*express*   一个代表 **SoftEdgeFormat** 对象的变量。

#### **SoftEdgeFormat.Type**

返回或设置一个 **MsoBevelType** 常量，该常量代表使用软边缘格式的图像边缘的棱台类型。可读写。

**语法**

**express.Type**

*express*   一个代表 **SoftEdgeFormat** 对象的变量。

适用环境：web

适用平台：windows/linux