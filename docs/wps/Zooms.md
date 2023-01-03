#### **Zooms**



[Zoom ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Zoom/Zoom%20.htm#jsObject_Zoom)对象的集合，该集合代表每个视图（如大纲视图、普通视图或页面视图）的缩放选项。

**方法**

|                                                              | 名称     | 说明                                 |
| ------------------------------------------------------------ | -------- | ------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回指定的 **Zoom** 对象。返回Zoom值 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Zooms** 对象的父对象。 |

**成员方法**

#### **Zooms.Item**

返回指定的 **Zoom** 对象。返回Zoom值

**语法**

**express.Item(Index)**

*express*   一个代表 **Zooms** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**   | **说明**             |
| -------- | ------------- | -------------- | -------------------- |
| *Index*  | 可选          | **WdViewType** | 指定的显示比例类型。 |

**成员属性**

#### **Zooms.Application**

返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **Zooms** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Zooms.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Zooms** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 Creator 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Zooms.Parent**

返回一个 **Object** 类型值，该值代表指定 **Zooms** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Zooms** 对象的变量。

适用环境：web

适用平台：windows/linux