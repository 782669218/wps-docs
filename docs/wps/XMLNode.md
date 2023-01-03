#### **XMLNode**



表示应用于文档的单个 XML 元素。

**说明**

每个已应用于文档的 XML 元素都显示为**“XML 结构”**任务窗格的树视图控件中的一个节点。树视图中的每个节点都是 **XMLNode** 对象的一个实例。树视图的层次结构表明节点是否包含子节点。

使用 **XMLNodes** 集合的 **Item** 方法可返回一个 **XMLNode** 对象。使用 **Validate** 方法可根据所应用的架构验证 XML 元素是否有效、必需的子元素是否存在并按要求的顺序显示。运行 **Validate** 方法后，使用 **ValidationStatus** 属性可验证元素是否有效，使用 **ValidationErrorText** 属性可显示关于如下内容的信息：用户应采用哪些措施来确保文档符合该 XML 架构的规则。

以下示例验证活动文档中的每个 XML 元素。如果根据架构发现元素无效，则该示例向用户返回一条消息解释问题所在。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let objNode = Application.ActiveDocument.XMLNodes     for(let i = 1; i <= objNode.Count; i++) {         objNode.Item(i).Validate()         if(objNode.ValidationStatus != wdXMLValidationStatusOK) {             alert(objNode.ValidationErrorText(true))         }     } }` |

**方法**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**               | 将指定的 XML 元素（不包括 XML 标记）复制到剪贴板。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**                | 将指定的 XML 元素从文档中移到剪贴板上。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**             | 删除 XML 文档中指定的 XML 元素。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RemoveChild**        | 从指定的元素中删除子元素。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectNodes**        | 返回一个 **XMLNodes** 集合，该集合代表与 XPath 参数匹配的所有子元素（它们的顺序与它们在指定的 XML 元素中出现的顺序相同）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectSingleNode**   | 返回一个 **XMLNode** 对象，该对象代表指定的 XML 元素中第一个与 XPath 参数匹配的子元素。返回XMLNode值 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetValidationError** | 更改向用户显示的指定节点的验证错误文本，并强制 WPS 将节点报告为无效。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Validate**           | 针对附加到文档的 XML 架构验证单个 XML 元素。                 |

**属性**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**          | 返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Attributes**           | 返回一个 **XMLNodes** 集合，该集合代表指定元素的属性。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseName**             | 返回一个 **String** 类型的值，该值代表不带任何前缀的元素名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChildNodeSuggestions** | 返回一个 **XMLChildNodeSuggestions** 集合，代表指定元素的子元素列表。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChildNodes**           | 返回一个 **XMLNodes** 集合，该集合代表指定元素的子元素。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**              | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstChild**           | 返回一个 **DiagramNode** 对象，该对象代表父节点的第一个子节点。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasChildNodes**        | 返回 **Boolean** 值，该值表示 XML 节点是否有子节点。只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LastChild**            | 返回一个 **XMLNode** 对象，该对象代表 XML 元素的最后一个子节点。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Level**                | 返回一个 [WdXMLNodeLevel ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdXMLNodeLevel%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表 XML 元素是段落的一部分、是整个段落、包含在表格单元格中还是包含表格行。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NamespaceURI**         | 返回一个 **String** 类型的值，该值代表指定对象的架构命名空间的统一资源标识符 (URI)。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NextSibling**          | 返回一个 **XMLNode** 对象，该对象代表文档中与指定元素同级的下一个元素。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NodeType**             | 返回一个 **WdXMLNodeType** 常量，该常量代表节点的类型。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NodeValue**            | 返回或设置一个 **String** 类型的值，该值代表 XML 元素的值。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OwnerDocument**        | 返回一个代表指定 XML 元素的父文档的 **Document** 对象。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**               | 返回一个 **Object** 类型值，该值代表指定 **XMLNode** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentNode**           | 返回一个代表指定元素的父元素的 **XMLNode** 对象。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PlaceholderText**      | 设置或返回一个 **String** 类型的值，该值代表为不包含文本的元素显示的文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreviousSibling**      | 返回一个 **XMLNode** 对象，该对象代表文档中与指定元素同级的前一个元素。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**                | 返回一个 **Range** 对象，该对象代表包含在指定对象中的文档部分。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**                 | 返回或设置 XML 元素中包含的文本。**String** 类型，可读写。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ValidationErrorText**  | 返回一个 **String** 类型的值，代表关于 **XMLNode** 对象验证错误的说明。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordOpenXML**          | 返回一个 **String** 类型的值，该值以 WPS Open XML 格式表示节点的 XML。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XML**                  | 返回 **String** 值，该值表示包含在 XML 节点中的文本（有或没有 XML 标记）。只读。 |

**成员方法**

#### **XMLNode.Copy**

将指定的 XML 元素（不包括 XML 标记）复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **XMLNode** 对象的变量。

#### **XMLNode.Cut**

将指定的 XML 元素从文档中移到剪贴板上。

**语法**

**express.Cut()**

*express*   一个代表 **XMLNode** 对象的变量。

#### **XMLNode.Delete**

删除 XML 文档中指定的 XML 元素。

**语法**

**express.Delete()**

*express*   一个代表 **XMLNode** 对象的变量。

#### **XMLNode.RemoveChild**

从指定的元素中删除子元素。

**语法**

**express.RemoveChild(ChildElement)**

*express*   一个代表 **XMLNode** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**         |
| -------------- | ------------- | ------------ | ---------------- |
| *ChildElement* | 必选          | **XMLNode**  | 要删除的子元素。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下示例删除活动文档中第一个元素的第一个子元素。 */ Application.ActiveDocument.XMLNodes.Item(1).RemoveChild(Application.ActiveDocument.XMLNodes.Item(1).ChildNodes.Item(1))` |

#### **XMLNode.SelectNodes**

返回一个 **XMLNodes** 集合，该集合代表与 XPath 参数匹配的所有子元素（它们的顺序与它们在指定的 XML 元素中出现的顺序相同）。

**语法**

**express.SelectNodes(XPath, PrefixMapping, FastSearchSkippingTextNodes)**

*express*   一个代表 **XMLNode** 对象的变量。

**参数**

| **名称**                      | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *XPath*                       | 必选          | **String**   | 指定有效的 XPath 字符串。                                    |
| *PrefixMapping*               | 可选          | **String**   | 提供用于搜索的架构中的前缀。如果 XPath 参数使用名称来搜索元素，则可使用 PrefixMapping 参数。 |
| *FastSearchSkippingTextNodes* | 可选          |              | 如果该参数值为 True，则搜索指定节点时跳过所有文本节点。如果该参数值为 False，则搜索时包含文本节点。默认值为 False。 |

#### **XMLNode.SelectSingleNode**

返回一个 **XMLNode** 对象，该对象代表指定的 XML 元素中第一个与 XPath 参数匹配的子元素。

返回XMLNode值

**语法**

**express.SelectSingleNode(XPath, PrefixMapping, FastSearchSkippingTextNodes)**

*express*   一个代表 **XMLNode** 对象的变量。

**参数**

| **名称**                      | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *XPath*                       | 必选          | **String**   | 指定有效的 XPath 字符串。                                    |
| *PrefixMapping*               | 可选          | **String**   | 提供用于搜索的架构中的前缀。如果 XPath 参数使用名称来搜索元素，则可使用 PrefixMapping 参数。 |
| *FastSearchSkippingTextNodes* | 可选          | **Boolean**  | 如果该参数值为 True，则搜索指定节点时跳过所有文本节点。如果该参数值为 False，则搜索时包含文本节点。默认值为 False。 |

#### **XMLNode.SetValidationError**

更改向用户显示的指定节点的验证错误文本，并强制 WPS 将节点报告为无效。

**语法**

**express.SetValidationError(Status, ErrorText, ClearedAutomatically)**

*express*   一个代表 **XMLNode** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型**              | **说明**                                                     |
| ---------------------- | ------------- | ------------------------- | ------------------------------------------------------------ |
| *Status*               | 必选          | **WdXMLValidationStatus** | 指定设置验证状态错误文本 (wdXMLValidationStatusCustom)，还是清除验证状态错误文本 (wdXMLValidationStatusOK)。 |
| *ErrorText*            | 可选          | **Variant**               | 显示给用户的文本。如果 Status 参数设置为 wdXMLValidationStatusOK，则将该参数保留为空。 |
| *ClearedAutomatically* | 可选          | **Boolean**               | 如果值为 True，则指定节点上出现下一验证事件时，将自动清除错误消息。如果值为 False，则需要运行 SetValidationError 方法（Status 参数为 wdXMLValidationStatusOK）来清除自定义错误文本。 |

**说明**

若要设置自定义错误文本，请使用 **wdXMLValidationStatusCustom** 常量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例指定自定义验证错误文本。*/ function test() {     let objNode = Application.ActiveDocument.XMLNodes.Item(1)     objNode.SetValidationError(wdXMLValidationStatusCustom,"Error Text",true) }` |

#### **XMLNode.Validate**

针对附加到文档的 XML 架构验证单个 XML 元素。

**语法**

**express.Validate()**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

使用具有 **ValidationStatus** 和 **ValidationErrorText** 属性的 **Validate** 方法，根据所应用的架构确定 XML 元素是否有效，并且确定向用户显示的错误文本内容。使用 **SetValidationError** 方法可以用自定义验证错误来替代架构冲突。

当您运行 **Validate** 方法时，WPS 会用具有验证错误的 XML 节点的集合填充 **Document** 对象的 **XMLSchemaViolations** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下示例检查活动文档中的每个元素和属性，并显示一条消息，其中包含根据架构没有通过验证的元素和属性以及原因说明*/ function test() {     let objNode = Application.ActiveDocument.XMLNodes     let strValid      for(let i=1; i <= objNode.Count; i++) {         objNode.Item(i).Validate()         if(objNode.Item(i).ValidationStatus != wdXMLValidationStatusOK) {             strValid = strValid + objNode.Item(i).BaseName + '\t'  + objNode.Item(i).ValidationErrorText + '\r\n'         }     }         alert("The following elements do not validate against " + "the schema." + '\r\n' + '\r\n' + strValid + '\r\n' + "You should fix these elements before continuing.") }` |

**成员属性**

#### **XMLNode.Application**

返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **XMLNode.Attributes**

返回一个 **XMLNodes** 集合，该集合代表指定元素的属性。

**语法**

**express.Attributes**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

使用 **Attributes** 属性返回的 **XMLNodes** 集合中的所有 **XMLNode** 对象的 **NodeType** 属性值均为 **wdXMLNodeAttribute**。

#### **XMLNode.BaseName**

返回一个 **String** 类型的值，该值代表不带任何前缀的元素名称。

**语法**

**express.BaseName**

*express*   一个代表 **XMLNode** 对象的变量。

#### **XMLNode.ChildNodeSuggestions**

返回一个 **XMLChildNodeSuggestions** 集合，代表指定元素的子元素列表。

**语法**

**express.ChildNodeSuggestions**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

**XMLChildNodeSuggestions** 集合中的每个 **XMLChildNodeSuggestion** 对象都是允许的、可能做为子元素的 XML 元素列表中的一项，该列表位于**“XML 结构”**任务窗格底部。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例循环遍历活动文档中选定的第一个元素的建议子元素，并在插入点插入所有允许的元素。*/ function test() {     let objSuggestion     let objNode = Application.Selection.XMLParentNode          for(objSuggestion = 1; objSuggestion <= objNode.ChildNodeSuggestions.Count; objSuggestion++) {         XMLChildNodeSuggestions.Item(objSuggestion).Insert()         Application.Selection.MoveRight()     } }` |

#### **XMLNode.ChildNodes**

返回一个 **XMLNodes** 集合，该集合代表指定元素的子元素。

**语法**

**express.ChildNodes**

*express*   一个代表 **XMLNode** 对象的变量。

#### **XMLNode.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **XMLNode.FirstChild**

返回一个 **DiagramNode** 对象，该对象代表父节点的第一个子节点。只读。

**语法**

**express.FirstChild**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

使用 **LastChild** 属性访问最后一个子节点。使用 **Root** 属性访问图表中的父节点。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一个组织图表添加至活动文档，添加三个节点，并将第一个和最后一个子节点赋给变量。*/ function test() {     let shpDiagram     let dgnRoot     let dgnFirstChild     let dgnLastChild     let intCount      //Add organizational chart diagram to the current document     shpDiagram = Application.ActiveDocument.Shapes.AddDiagram(msoDiagramOrgChart, 10, 15, 400, 475)      //Add the first node to the diagram     dgnRoot = shpDiagram.DiagramNode.Children.AddNode()      //Add three child nodes     for(intCount = 1; intCount <= 3; intCount++) {         dgnRoot.Children.AddNode()     }      //Assign the first and last child nodes to variables     dgnFirstChild = dgnRoot.Children.FirstChild     dgnLastChild = dgnRoot.Children.LastChild }` |

#### **XMLNode.HasChildNodes**

返回 **Boolean** 值，该值表示 XML 节点是否有子节点。只读。

**语法**

**express.HasChildNodes**

*express*   一个代表 **XMLNode** 对象的变量。

#### **XMLNode.LastChild**

返回一个 **XMLNode** 对象，该对象代表 XML 元素的最后一个子节点。

**语法**

**express.LastChild**

*express*   一个代表 **XMLNode** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例访问活动文档中第二个元素的最后一个子元素。*/ let objNode = Application.ActiveDocument.XMLNodes.Item(2).LastChild` |

#### **XMLNode.Level**

返回一个 [WdXMLNodeLevel ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdXMLNodeLevel%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表 XML 元素是段落的一部分、是整个段落、包含在表格单元格中还是包含表格行。只读。

**语法**

**express.Level**

*express*   一个代表 **XMLNode** 对象的变量。

#### **XMLNode.NamespaceURI**

返回一个 **String** 类型的值，该值代表指定对象的架构命名空间的统一资源标识符 (URI)。只读。

**语法**

**express.NamespaceURI**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释         |
| ------------------------------------------------------------ |
| 如果创建的是用于 WPS 的 XML 架构，则强烈建议在架构中指定 targetNamespace 设置。 |

#### **XMLNode.NextSibling**

返回一个 **XMLNode** 对象，该对象代表文档中与指定元素同级的下一个元素。

**语法**

**express.NextSibling**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

如果指定元素是 **XMLNodes** 集合中的最后一个元素，则此属性返回 **Nothing**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回活动文档中第二个元素的下一个同级元素。*/ let objNode = Application.ActiveDocument.XMLNodes.Item(2).NextSibling` |

#### **XMLNode.NodeType**

返回一个 **WdXMLNodeType** 常量，该常量代表节点的类型。

**语法**

**express.NodeType**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

**XMLNode** 对象可以是 XML 元素或元素的属性。请使用 **NodeType** 属性确定要使用的节点的类型，以便您不会试图对该节点执行无效操作。例如，虽然 **Attributes** 属性出现在 **XMLNode** 对象的可用属性列表中，但它仅适用于元素节点。

#### **XMLNode.NodeValue**

返回或设置一个 **String** 类型的值，该值代表 XML 元素的值。可读写。

**语法**

**express.NodeValue**

*express*   一个代表 **XMLNode** 对象的变量。

#### **XMLNode.OwnerDocument**

返回一个代表指定 XML 元素的父文档的 **Document** 对象。

**语法**

**express.OwnerDocument**

*express*   一个代表 **XMLNode** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例访问活动选定内容中第一个 XML 元素的父文档。*/ let objDoc = Application.Selection.XMLNodes.Item(1).OwnerDocument` |

#### **XMLNode.Parent**

返回一个 **Object** 类型值，该值代表指定 **XMLNode** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **XMLNode** 对象的变量。

#### **XMLNode.ParentNode**

返回一个代表指定元素的父元素的 **XMLNode** 对象。

**语法**

**express.ParentNode**

*express*   一个代表 **XMLNode** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例访问活动文档中选定文本的 XML 父节点。*/ let objNode = Application.Selection.XMLParentNode.ParentNode` |

#### **XMLNode.PlaceholderText**

设置或返回一个 **String** 类型的值，该值代表为不包含文本的元素显示的文本。

**语法**

**express.PlaceholderText**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

仅当清除了**“XML 结构”**任务窗格中的**“在文档中显示 XML 标记”**复选框时，才在 WPS 中显示占位符文本。**“在文档中显示 XML 标记”**复选框对应于 **ShowXMLMarkup** 属性。

#### **XMLNode.PreviousSibling**

返回一个 **XMLNode** 对象，该对象代表文档中与指定元素同级的前一个元素。

**语法**

**express.PreviousSibling**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

如果指定元素是 **XMLNodes** 集合中的第一个元素，则此属性返回 **Nothing**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回活动文档中第三个元素的前一个同级元素。*/ let objNode = Application.ActiveDocument.XMLNodes.Item(3).PreviousSibling` |

#### **XMLNode.Range**

返回一个 **Range** 对象，该对象代表包含在指定对象中的文档部分。只读。

**语法**

**express.Range**

*express*   一个代表 **XMLNode** 对象的变量。

#### **XMLNode.Text**

返回或设置 XML 元素中包含的文本。**String** 类型，可读写。

**语法**

**express.Text**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

**Text** 属性返回 XML 元素中的无格式纯文本。设置该属性，可替换现有文本。

#### **XMLNode.ValidationErrorText**

返回一个 **String** 类型的值，代表关于 **XMLNode** 对象验证错误的说明。

**语法**

**express.ValidationErrorText**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

**ValidationErrorText(Boolean Advanced),Advanced代表显示的错误文字是验证错误说明的高级版本，该版本来自 WPS 所包含的 MSXML 5.0 组件。**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例检查活动文档中的每个元素，并显示一条消息，其中包含未通过架构验证的元素和属性，并说明相关原因。*/ function test() {     let objNode = Application.ActiveDocument.XMLNodes     let strValid      for(i = 1; i <= objNode.Count; i++) {          objNode.Item(i).Validate()         if(objNode.Item(i).ValidationStatus != wdXMLValidationStatusOK) {             strValid = strValid + objNode.Item(i).BaseName + "\t" + objNode.Item(i).ValidationErrorText + "\r\n"         }     }      alert("The following elements do not validate against " + "the schema." + "\r\n" + "\r\n" + strValid + "\r\n" + "You should fix these elements before continuing.") } ` |

#### **XMLNode.WordOpenXML**

返回一个 **String** 类型的值，该值以 WPS Open XML 格式表示节点的 XML。只读。

**语法**

**express.WordOpenXML**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

此属性只返回文档中为表示指定的 XML 节点所需的 XML。

#### **XMLNode.XML**

返回 **String** 值，该值表示包含在 XML 节点中的文本（有或没有 XML 标记）。只读。

**语法**

**express.XML**

*express*   一个代表 **XMLNode** 对象的变量。

**说明**

参数

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *DataOnly* | 可选          | **Boolean**  | 指定返回的 XML 是否带标记。如果为 **True**，将返回 XML 节点中包含的文本，但不带 XML 标记。如果为 **False**，将返回带 XML 标记的文本。 |

适用环境：web

适用平台：windows/linux