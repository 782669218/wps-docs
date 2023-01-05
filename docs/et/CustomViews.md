**CustomViews**



由自定义工作簿视图组成的集合。

**说明**

每个视图由一个 **CustomView** 对象代表。

使用 **CustomViews** 属性可返回 **CustomViews** 集合。使用 **Add** 方法可新建自定义视图，并将它添加到 **CustomViews** 集合中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列新建一个名为“Summary”的自定义视图。*/ Application.ActiveWorkbook.CustomViews.Add("Summary", true, true)` |

**方法**

|                                                              | 名称     | 说明                   |
| ------------------------------------------------------------ | -------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 新建自定义视图。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **CustomViews.Add**

新建自定义视图。

**语法**

**express.Add(ViewName, PrintSettings, RowColSettings)**

*express*   一个代表 **CustomViews** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ViewName*       | 必选          | **String**   | 新视图的名称。                                               |
| *PrintSettings*  | 可选          | **Variant**  | 如果为 True，则在自定义视图中包括打印设置。                  |
| *RowColSettings* | 可选          | **Variant**  | 如果为 True，则在自定义视图中包括隐藏行和隐藏列（包括筛选信息）的设置。 |

**返回值**

一个 CustomView 对象，它代表新的自定义视图。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在活动工作簿中新建一个自定义视图，并命名为“Summary”。*/ Application.ActiveWorkbook.CustomViews.Add("Summary", true, true)` |

#### **CustomViews.Item**

从集合中返回一个对象。

**语法**

**express.Item(ViewName)**

*express*   一个代表 **CustomViews** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**             |
| ---------- | ------------- | ------------ | -------------------- |
| *ViewName* | 必选          | **Variant**  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 CustomView 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在判断名为“Current Inventory”的自定义视图中是否包含打印设置。*/ alert(Application.ActiveWorkbook.CustomViews.Item("Current Inventory").PrintSettings == true)  ` |

**成员属性**

#### **CustomViews.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CustomViews** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ let myObject = Application.ActiveWorkbook if(myObject.Application.Value == "ET") {     alert("This is an ET Application object.") } else {     alert("This is not an ET Application object.") } }` |

#### **CustomViews.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **CustomViews** 对象的变量。

#### **CustomViews.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **CustomViews** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **CustomViews.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CustomViews** 对象的变量。

适用环境：web

适用平台：windows/linux