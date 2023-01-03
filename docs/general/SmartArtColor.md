#### **SmartArtColor**



为 SmartArt 图表选择配色方案。

**说明**

模拟 WPS Office Fluent 功能区用户界面中“SmartArt 工具”选项卡上“设计”组内的“更改颜色”命令上的命令。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码设置 Smart Art 图表的配色方案。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Color = Application.SmartArtColors.Item(1)` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，该对象代表 **SmartArtColor** 对象的容器应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Category**    | 检索与 SmartArt 颜色样式关联的主类别名称。只读。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建了此对象的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Description** | 检索 SmartArt 颜色样式的说明。只读。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**          | 检索关联的 SmartArt 颜色样式的唯一 Id。只读。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 检索 SmartArt 颜色样式的字符串名称。只读。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回调用对象。只读。                                         |

**成员属性**

#### **SmartArtColor.Application**

获取一个 **Application** 对象，该对象代表 **SmartArtColor** 对象的容器应用程序。只读。

**语法**

**express.Application**

*express*   一个代表 **SmartArtColor** 对象的变量。

#### **SmartArtColor.Category**

检索与 SmartArt 颜色样式关联的主类别名称。只读。

**语法**

**express.Category**

*express*   一个代表 **SmartArtColor** 对象的变量。

#### **SmartArtColor.Creator**

返回一个 32 位整数，该整数指示在其中创建了此对象的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **SmartArtColor** 对象的变量。

#### **SmartArtColor.Description**

检索 SmartArt 颜色样式的说明。只读。

**语法**

**express.Description**

*express*   一个代表 **SmartArtColor** 对象的变量。

#### **SmartArtColor.Id**

检索关联的 SmartArt 颜色样式的唯一 Id。只读。

**语法**

**express.Id**

*express*   一个代表 **SmartArtColor** 对象的变量。

**说明**

与此属性关联的 ID 区分大小写。

#### **SmartArtColor.Name**

检索 SmartArt 颜色样式的字符串名称。只读。

**语法**

**express.Name**

*express*   一个代表 **SmartArtColor** 对象的变量。

#### **SmartArtColor.Parent**

返回调用对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **SmartArtColor** 对象的变量。

适用环境：web

适用平台：windows/linux