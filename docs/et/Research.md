**Research**



代表 **Research** 查询的控件。

**说明**

使用 **Research** 查询时，必须具有对应于实时数据源的现有 GUID。如果相应数据源不可用或不存在，将发生运行时错误。

**方法**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IsResearchService** | 指示 *ServiceID* 参数中指定的 GUID 是否对应于当前配置的服务。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Query**             | 指定搜索查询。                                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetLanguagePair**   | 设置翻译服务的语言。                                         |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **Research.IsResearchService**

指示 *ServiceID* 参数中指定的 GUID 是否对应于当前配置的服务。

**语法**

**express.IsResearchService(ServiceID)**

*express*   一个代表 **Research** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                      |
| ----------- | ------------- | ------------ | ----------------------------- |
| *ServiceID* | 必选          | **String**   | 指定一个标识搜索服务的 GUID。 |

**返回值**

Boolean

#### **Research.Query**

指定搜索查询。

**语法**

**express.Query(ServiceID, QueryString, QueryLanguage, UseSelection, RequeryContextXML, NewQueryContextXML, LaunchQuery)**

*express*   一个代表 **Research** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ServiceID*          | 必选          | **String**   | 指定一个标识搜索服务的 GUID。                                |
| *QueryString*        | 可选          | **Variant**  | 指定查询字符串。                                             |
| *QueryLanguage*      | 可选          | **Variant**  | 指定查询字符串的查询语言。                                   |
| *UseSelection*       | 可选          | **Variant**  | 指定查询字符串的查询语言。                                   |
| *RequeryContextXML*  | 可选          | **Variant**  | 指定包含所请求内容的 xml 文件。                              |
| *NewQueryContextXML* | 可选          | **Variant**  | 指定包含新查询内容的 XML 文件。                              |
| *LaunchQuery*        | 可选          | **Boolean**  | 如果为 True，则启动查询。如果为 False，则显示“信息检索”任务窗格，该任务窗格用于搜索指定的信息检索服务。 |

**返回值**

Variant

#### **Research.SetLanguagePair**

设置翻译服务的语言。

**语法**

**express.SetLanguagePair(LanguageFrom, LanguageTo)**

*express*   一个代表 **Research** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**             |
| -------------- | ------------- | ------------ | -------------------- |
| *LanguageFrom* | 必选          | **Long**     | 指定翻译的源语言。   |
| *LanguageTo*   | 必选          | **Long**     | 指定翻译的目标语言。 |

**返回值**

Variant

**成员属性**

#### **Research.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Research** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **Research.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Research** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Research.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Research** 对象的变量。

适用环境：web

适用平台：windows/linux