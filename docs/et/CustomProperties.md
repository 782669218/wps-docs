**CustomProperties**



由代表附加信息的 **CustomProperty** 对象组成的集合，这些信息可用作 XML 的元数据。

**说明**

使用 **Worksheet** 对象的 **CustomProperties** 属性返回 **CustomProperties** 集合。

返回 **CustomProperties** 集合后，可根据选择向工作表和智能标记中添加元数据。

若要向工作表添加元数据，请在 **Add** 方法中使用 **CustomProperties** 属性。

下例演示了该功能。在此示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let wksSheet1 = Application.ActiveSheet      // Add metadata to worksheet.     wksSheet1.CustomProperties.Add("Market", "Nasdaq")      // Display metadata.     let cusProperties = wksSheet1.CustomProperties.Item(1)     alert(cusProperties.Name + "\t" + cusProperties.Value) }` |

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 添加自定义属性信息，返回一个代表自定义属性信息的 **CustomProperty**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/CustomProperty/CustomProperty%20.htm#jsObject_CustomProperty)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。                                       |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **CustomProperties.Add**

添加自定义属性信息，返回一个代表自定义属性信息的 **CustomProperty**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/CustomProperty/CustomProperty%20.htm#jsObject_CustomProperty)对象。

**语法**

**express.Add(Name, Value)**

*express*   一个代表 **CustomProperties** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**           |
| -------- | ------------- | ------------ | ------------------ |
| *Name*   | 必选          | **String**   | 自定义属性的名称。 |
| *Value*  | 必选          | **Variant**  | 自定义属性的值。   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例向活动工作表添加标识符信息，并将其名称和值返回给用户*/ function test() {      let wksSheet1 = Application.ActiveSheet      // Add metadata to worksheet.     wksSheet1.CustomProperties.Add("Market", "Nasdaq")      // Display metadata.     let cusProperties = wksSheet1.CustomProperties.Item(1)     alert(cusProperties.Name + "\t" + cusProperties.Value)  }` |

#### **CustomProperties.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **CustomProperties** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 下例演示了该功能。在此示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值*/ function test() {      let wksSheet1 = Application.ActiveSheet      // Add metadata to worksheet.     wksSheet1.CustomProperties.Add("Market", "Nasdaq")      // Display metadata.     let cusProperties = wksSheet1.CustomProperties.Item(1)         alert(cusProperties.Name + "\t" + cusProperties.Value)  }` |

**成员属性**

#### **CustomProperties.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CustomProperties** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ function test(){ 	let myObject = Application.ActiveWorkbook 	if(myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} 	else { 		alert("This is not an ET Application object.") 	} }` |

#### **CustomProperties.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **CustomProperties** 对象的变量。

#### **CustomProperties.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **CustomProperties** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **CustomProperties.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CustomProperties** 对象的变量。

适用环境：web

适用平台：windows/linux