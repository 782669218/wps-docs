**RecentFiles**



代表最近使用过的文件的列表。

**说明**

每个文件都由一个 **RecentFile** 对象代表。

**方法**

|                                                              | 名称    | 说明                             |
| ------------------------------------------------------------ | ------- | -------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add** | 向最近使用的文件列表中添加文件。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 从集合中返回一个对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Maximum**     | 返回或设置最近使用文件清单中文件数目的上限。可为 0（零）到 9 之间的数字。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **RecentFiles.Add**

向最近使用的文件列表中添加文件。

**语法**

**express.Add(Name)**

*express*   一个代表 **RecentFiles** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明** |
| -------- | ------------- | ------------ | -------- |
| *Name*   | 必选          | **String**   | 文件名。 |

**返回值**

包含在集合中的一个 RecentFile 对象。

**示例**

此示例将文件 Oscar.xls 添加到最近使用的文件列表中。

| 示例代码复制                               |
| ------------------------------------------ |
| `Application.RecentFiles.Add("Oscar.xls")` |

**成员属性**

#### **RecentFiles.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **RecentFiles** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook     if(myObject.Application.Value == "ET") { 　　　　    MsgBox("This is an ET Application object.") 　　　　} 　　　　else { 　　　　    MsgBox("This is not an ET Application object.") 　　　　} }` |

#### **RecentFiles.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **RecentFiles** 对象的变量。

#### **RecentFiles.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **RecentFiles** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **RecentFiles.Item**

从集合中返回一个对象。

**语法**

**express.Item**

*express*   一个代表 **RecentFiles** 对象的变量。

**示例**

此示例打开最近使用过的文件列表中的第二个文件。

| 示例代码复制                                     |
| ------------------------------------------------ |
| `此示例打开最近使用过的文件列表中的第二个文件。` |

#### **RecentFiles.Maximum**

返回或设置最近使用文件清单中文件数目的上限。可为 0（零）到 9 之间的数字。**Long** 类型，可读写。

**语法**

**express.Maximum**

*express*   一个代表 **RecentFiles** 对象的变量。

**示例**

本示例将最近使用的文件列表中的最多文件数设置为 6。

| 示例代码复制                          |
| ------------------------------------- |
| `Application.RecentFiles.Maximum = 6` |

#### **RecentFiles.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **RecentFiles** 对象的变量。

适用环境：web

适用平台：windows/linux