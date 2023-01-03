#### **XMLMapping**



表示 **ContentControl** 对象上自定义 XML 与内容控件之间的 XML 映射。XML 映射是内容控件中的文本与此文档的自定义 XML 数据存储中的 XML 元素之间的链接。

**说明**

使用**SetMapping**方法可以添加或更改使用 XPath 字符串的内容控件的 XML 映射。以下示例设置文档作者这一内置文档属性，在活动文档中插入新的内容控件，然后将此控件的 XML 映射设置为该内置文档属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.ActiveDocument.BuiltInDocumentProperties.Item("Author").Value = "David";     let objcc = Application.ActiveDocument.ContentControls.Add(wdContentControlDate, Application.ActiveDocument.Paragraphs.Item(1).Range);     let objMap = objcc.XMLMapping;     let bInMap = objMap.SetMapping("/ns1:coreProperties[1]/ns0:createdate[1]"); } ` |

SetMappingByNode

CustomXMLNode

SetMappingByNode

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.ActiveDocument.BuiltInDocumentProperties.Item("Author").Value = "David Jaffe"     let rng = Application.ActiveDocument.Paragraphs.Item(1).Range     let objcc = Application.ActiveDocument.ContentControls.Add(wdContentControlDate, rng)      let npURL = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"     let objNP =  Application.ActiveDocument.CustomXMLParts.SelectByNamespace(npURL).Item(1)     let objNode = objNP.DocumentElement.ChildNodes.Item(1)      let objMap = objcc.XMLMapping     let blnMap = objMap.SetMappingByNode(objNode) } ` |

CustomXMLPart

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	objCustomPart = Application.ActiveDocument.CustomXMLParts.Add(); 	objCustomPart.LoadXML ("<books><book><author>Matt Hink</author>" +             "<title>Migration Paths of the Red Breasted Robin</title>" +             "<genre>non-fiction</genre><price>29.95</price>" +             "<pub_date>2/1/2007</pub_date><abstract>You see them in " +             "the spring outside your windows.  You hear their lovely " +             "songs wafting in the warm spring air.  Now follow the path " +             "of the red breasted robin as it migrates to warmer climes " +             "in the fall, and then back to your back yard in the spring." +             "</abstract></book></books>")  	Application.ActiveDocument.Range().InsertParagraphBefore() 	objRange = Application.ActiveDocument.Paragraphs.Item(1).Range 	objCustomNode = objCustomPart.SelectSingleNode("/books/book/title") 	objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange) 	objCustomControl.XMLMapping.SetMappingByNode(objCustomNode) 	 	objRange.InsertParagraphAfter() 	objRange = Application.ActiveDocument.Paragraphs.Item(2).Range 	objCustomNode = objCustomPart.SelectSingleNode("/books/book/abstract") 	objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange) 	objCustomControl.XMLMapping.SetMappingByNode(objCustomNode) 	 	alert(objCustomControl.XMLMapping.IsMapped) }` |

Delete 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.ContentControls.Item(1).XMLMapping.Delete()` |

IsMapped 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.ContentControls.Item(1).XMLMapping.IsMapped` |

Custom

XMLNode 

CustomXMLPart 

CustomXMLNode

CustomXMLPart

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**           | 从父内容控件中删除 XML 映射。                                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetMapping**       | 允许创建或更改内容控件上的 XML 映射。如果 WPS 将内容控件映射到文档的自定义 XML 数据存储中的自定义 XML 节点，将返回 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetMappingByNode** | 允许创建或更改内容控件上的 XML 数据映射。如果 WPS 将内容控件映射到文档的自定义 XML 数据存储中的自定义 XML 节点，将返回 **True**。 |

**属性**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**    | 返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象代表 WPS 应用程序。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**        | 返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomXMLNode**  | 返回 **CustomXMLNode** 对象，该对象表示文档中的内容控件映射到的自定义 XML 节点（位于数据存储中）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomXMLPart**  | 返回 **CustomXMLPart** 对象，该对象表示文档中的内容控件映射到的自定义 XML 部件。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsMapped**       | 返回 **Boolean** 值，该值表示文档中的内容控件是否映射到文档的 XML 数据存储内的 XML 节点。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**         | 返回一个 **Object** 类型的值，该值代表指定 **XMLMapping** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrefixMappings** | 返回 **String** 值，该值表示用于为当前 XML 映射计算 XPath 的前缀映射。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XPath**          | 返回 **String** 值，该值表示 **XML** 映射的 **XPath**，其计算结果是当前映射的 **XML** 节点。只读。 |

**成员方法**

#### **XMLMapping.Delete**

从父内容控件中删除 XML 映射。

**语法**

**express.Delete()**

*express*   一个代表 **XMLMapping** 对象的变量。

**说明**

此操作删除 XML 映射。XML 数据和内容控件仍保留在文档中。

#### **XMLMapping.SetMapping**

允许创建或更改内容控件上的 XML 映射。如果 WPS 将内容控件映射到文档的自定义 XML 数据存储中的自定义 XML 节点，将返回 **True**。

**语法**

**express.SetMapping(XPath, PrefixMapping, Source)**

*express*   一个代表 **XMLMapping** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**      | **说明**                                                     |
| --------------- | ------------- | ----------------- | ------------------------------------------------------------ |
| *XPath*         | 必选          | **String**        | 指定 XPath 字符串，该字符串表示内容控件所映射到的 XML 节点。无效的 XPath 字符串将导致运行时错误。 |
| *PrefixMapping* | 可选          | **String**        | 指定在查询 XPath 参数所提供的表达式时要使用的前缀映射。如果省略， WPS 将使用当前文档中指定的自定义 XML 部件的前缀映射集。 |
| *Source*        | 可选          | **CustomXMLPart** | 指定内容控件要映射到的自定义 XML 数据。如果省略此参数，将依据当前文档中的所有自定义 XML 来计算 XPath，并且用 XPath 在其中解析为 XML 节点的第一个 CustomXMLPart 来建立映射。 |

**说明**

如果 XML 映射已经存在， WPS 将替换现有 XML 映射，并以新映射的 XML 节点的内容替换内容控件的文本。如果指定的 XPath 的计算结果不是指定的自定义 XML 部件中的 XML 节点，那么仍可以指定映射，并且将创建一个映射。当指定的 XPath 的计算结果为指定的自定义 XML 部件中的 XML 节点时，此映射将自动链接。

请参阅 [SetMappingByNode ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/XMLMapping/XMLMapping%20.htm#XMLMapping.SetMappingByNode)方法。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释 |
| ---------------------------------------------------- |
| 为格式文本内容控件创建映射将导致运行时错误。         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下示例插入自定义 XML 部件并为该自定义部件设置 XML，然后在文档开头插入两个内容控件，并将控件的内容映射为自定义部件中 XML 元素的内容*/ function test() {     let objRange     let objCustomPart     let objCustomControl      objCustomPart = Application.ActiveDocument.CustomXMLParts.Add()     objCustomPart.LoadXML ("<books><book><author>Matt Hink</author>" +         "<title>Migration Paths of the Red Breasted Robin</title>" +         "<genre>non-fiction</genre><price>29.95</price>" +         "<pub_date>2/1/2007</pub_date><abstract>You see them in " +         "the spring outside your windows.  You hear their lovely " +         "songs wafting in the warm spring air.  Now follow the path "+         "of the red breasted robin as it migrates to warmer climes " +         "in the fall, and then back to your back yard in the spring." +         "</abstract></book></books>")      Application.ActiveDocument.Range().InsertParagraphBefore()     objRange = Application.ActiveDocument.Paragraphs.Item(1).Range     objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange)     objCustomControl.XMLMapping.SetMapping("/books/book/title", null, objCustomPart)      objRange.InsertParagraphAfter()     objRange = Application.ActiveDocument.Paragraphs.Item(2).Range     objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange)     objCustomControl.XMLMapping.SetMapping("/books/book/abstract", null, objCustomPart) }` |

#### **XMLMapping.SetMappingByNode**

允许创建或更改内容控件上的 XML 数据映射。如果 WPS 将内容控件映射到文档的自定义 XML 数据存储中的自定义 XML 节点，将返回 **True**。

**语法**

**express.SetMappingByNode(Node)**

*express*   一个代表 **XMLMapping** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**      | **说明**                              |
| -------- | ------------- | ----------------- | ------------------------------------- |
| *Node*   | 必选          | **CustomXMLNode** | 指定当前内容控件所映射到的 XML 节点。 |

**说明**

如果 XML 映射已经存在，则 WPS 将替换现有 XML 映射，并以新映射的 XML 节点的内容替换内容控件的文本。请参阅**SetMapping**方法。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释 |
| ---------------------------------------------------- |
| 为格式文本内容控件创建映射将导致运行时错误。         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下示例设置文档作者这一内置文档属性，在活动文档中插入新的内容控件，然后将此控件的 XML 映射设置为该内置文档属性*/ function test() {     let objcc     let objNode     let objMap     let blnMap     let objNP      Application.ActiveDocument.BuiltInDocumentProperties.Item("Author").Value = "David Jaffe"      objcc = Application.ActiveDocument.ContentControls.Add(wdContentControlDate, ActiveDocument.Paragraphs.Item(1).Range)     objNP = Application.ActiveDocument.CustomXMLParts.SelectByNamespace("http://schemas.openxmlformats.org/package/2006/metadata/core-properties").Item(1)     objNode = objNP.DocumentElement.ChildNodes.Item(1)      objMap = objcc.XMLMapping     blnMap = objMap.SetMappingByNode(objNode)      /* 以下示例创建自定义 XML 部件，然后创建两个内容控件，并将每个内容控件映射到自定义 XML 中的特定节点*/     let objRange     let objCustomPart     let objCustomControl     let objCustomNode      objCustomPart = Application.ActiveDocument.CustomXMLParts.Add()     objCustomPart.LoadXML ("<books><book><author>Matt Hink</author>" +         "<title>Migration Paths of the Red Breasted Robin</title><genre>non-fiction</genre>" +         "<price>29.95</price><pub_date>2007-02-01</pub_date><abstract>" +         "You see them in the spring outside your windows.  You hear their lovely " +         "songs wafting in the warm spring air.  Now follow the path of the red breasted robin " +         "as it migrates to warmer climes in the fall, and then back to your back yard " +         "in the spring.</abstract></book></books>")      Application.ActiveDocument.Range().InsertParagraphBefore()     objRange = Application.ActiveDocument.Paragraphs.Item(1).Range     objCustomNode = objCustomPart.SelectSingleNode("/books/book/title")     objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange)     objCustomControl.XMLMapping.SetMappingByNode(objCustomNode)      objRange.InsertParagraphAfter()     objRange = Application.ActiveDocument.Paragraphs.Item(2).Range     objCustomNode = objCustomPart.SelectSingleNode("/books/book/abstract")     objCustomControl = Application.ActiveDocument.ContentControls.Add(wdContentControlText, objRange)     objCustomControl.XMLMapping.SetMappingByNode(objCustomNode) }` |

**成员属性**

#### **XMLMapping.Application**

返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **XMLMapping** 对象的变量。

#### **XMLMapping.Creator**

返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **XMLMapping** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释 |
| ---------------------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。              |

#### **XMLMapping.CustomXMLNode**

返回 **CustomXMLNode** 对象，该对象表示文档中的内容控件映射到的自定义 XML 节点（位于数据存储中）。

**语法**

**express.CustomXMLNode**

*express*   一个代表 **XMLMapping** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中插入新的内容控件和自定义 XML 部件，将内容控件映射到自定义 XML 部件中的某个节点，然后设置映射的 XML 节点的值。*/ function test() {     let objCC     let objPart     let objNode     let objMap      objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlText)     objPart = Application.ActiveDocument.CustomXMLParts.Add("<books><book>" +         "<author></author><title></title><genre></genre><price></price>" +         "<pub_date></pub_date><abstract></abstract></book></books>")      objMap = objCC.XMLMapping     objMap.SetMapping("/books/book/author", null, objPart)      objNode = objMap.CustomXMLNode     objNode.Text = "Matt Hink"      objCC.Range.Text = objNode.Text }` |

#### **XMLMapping.CustomXMLPart**

返回 **CustomXMLPart** 对象，该对象表示文档中的内容控件映射到的自定义 XML 部件。

**语法**

**express.CustomXMLPart**

*express*   一个代表 **XMLMapping** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下示例访问活动文档中的第一个内容控件以及它映射到的自定义 XML 部件，然后设置自定义 XML 部件中包含的某个 XML 节点的文本值*/ /* 此示例假定活动文档中至少有一个内容控件映射到包含 XPath 字符串指定的 XML 节点的自定义 XML 部件*/ function test() {     let objCC     let objPart     let objNode      objCC = Application.ActiveDocument.ContentControls.Item(1)     objPart = objCC.XMLMapping.CustomXMLPart     objNode = objPart.SelectSingleNode("/books/book/title")     objNode.Text = "Mystery of the Empty Chair" }` |

#### **XMLMapping.IsMapped**

返回 **Boolean** 值，该值表示文档中的内容控件是否映射到文档的 XML 数据存储内的 XML 节点。只读。

**语法**

**express.IsMapped**

*express*   一个代表 **XMLMapping** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下示例删除活动文档内当前映射的所有内容控件的 XML 映射*/ function test() {     let objCC = Application.ActiveDocument.ContentControls      for(let i = 1; i <= objCC.Count; i++){         if(objCC.Item(i).XMLMapping.IsMapped == true){             objCC.XMLMapping.Delete()         }     } }` |

#### **XMLMapping.Parent**

返回一个 **Object** 类型的值，该值代表指定 **XMLMapping** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **XMLMapping** 对象的变量。

#### **XMLMapping.PrefixMappings**

返回 **String** 值，该值表示用于为当前 XML 映射计算 XPath 的前缀映射。只读。

**语法**

**express.PrefixMappings**

*express*   一个代表 **XMLMapping** 对象的变量。

**说明**

若要为内容控件设置映射，请使用 **SetMapping** 方法或 **SetMappingByNode** 方法。

**示例**

| 示例代码复制                                       |
| -------------------------------------------------- |
| `请使用 SetMapping 方法或 SetMappingByNode 方法。` |

#### **XMLMapping.XPath**

返回 **String** 值，该值表示 **XML** 映射的 **XPath**，其计算结果是当前映射的 **XML** 节点。只读。

**语法**

**express.XPath**

*express*   一个代表 **XMLMapping** 对象的变量。

**说明**

若要为内容控件设置映射，请使用 [SetMapping ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/XMLMapping/XMLMapping%20.htm#XMLMapping.SetMapping)方法或 [SetMappingByNode ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/XMLMapping/XMLMapping%20.htm#XMLMapping.SetMappingByNode)方法。如果映射不是活动的，则使用此属性将返回错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下示例检查活动文档中第一个内容控件是否是日期控件，以及 XPath 字符串是否设置为特定的内置文档属性。如果 XPath 不匹配并且控件是日期控件，它将设置到该控件的映射*/ function test() {     let blnMap     let objCC = Application.ActiveDocument.ContentControls.Item(1)     let objMap = objCC.XMLMapping     if((objCC.Type == wdContentControlDate) && (objMap.XPath != "/ns1:coreProperties[1]/ns0:createdate[1]")) {         blnMap = objMap.SetMapping("/ns1:coreProperties[1]/ns0:createdate[1]")         if(blnMap == false) {             alert("Unable to map the content control.")         }     } }` |

适用环境：web

适用平台：windows/linux