#### **GroupShapes**



代表组合形状中的单个形状。组合形状中包含的每个形状都由一个 **Shape** 对象表示。

**方法**

|                                                              | 名称      | 说明                              |
| ------------------------------------------------------------ | --------- | --------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**  | 返回集合中的单个 **Shape** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Range** | 返回一个 **ShapeRange** 对象。    |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中图形的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **GroupShapes** 对象的父对象。 |

**成员方法**

#### **GroupShapes.Item**

返回集合中的单个 **Shape** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **GroupShapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

#### **GroupShapes.Range**

返回一个 **ShapeRange** 对象。

**语法**

**express.Range(Index)**

*express*   一个代表 **GroupShapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 指定要在指定区域内包含哪些形状。可以是指定 Shapes 集合内某个形状的索引号的整数，也可以是指定形状名称的字符串，或者是包含整数或字符串的数组。 |

**说明**

**ShapeRange** 对象不包含 **InlineShape** 对象。一个 **InlineShape** 对象相当于一个字符，并像字符一样位于文本区域中。**Shape** 对象定位到文本区域（默认情况下为所选内容），但对象可以位于页面上的任何位置。**Shape** 对象总是与定位的区域位于同一页上。

对 **Shape** 对象进行的大部分操作也适用于只包含一个形状的 **ShapeRange** 对象。有些操作如果用于包含多个形状的 **ShapeRange** 对象，则会产生错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个形状的填充前景色设置为紫色。*/ function test(){     let fill = Application.ActiveDocument.Shapes.Range(1).Fill     fill.ForeColor.RGB = (255, 0, 255)     fill.Visible = msoTrue }  /*以下示例为活动文档中变量形状应用阴影。*/ function test(strShpName){     Application.ActiveDocument.Shapes.Range(strShpName).Shadow.Type = msoShadow6 }  /*要调用前一个子程序，请在标准代码模块中输入以下代码。*/ function test(){     let shpArrow = Application.ActiveDocument.Shapes.AddShape(msoShapeLeftArrow,200,400,50,75)     shpArrow.Name = "myShape"     ShpRange2(shpArrow.Name) }  /*以下示例选择活动文档的第一和第三个形状。*/ function test(){     Application.ActiveDocument.Shapes.Range([1, 3]).Select() }  /*以下示例选择并删除活动文档的第一个形状中的形状。本示例假定第一个形状是画布形状。*/ function CanvasShapeRange(){     let rngCanvasShapes = Application.ActiveDocument.Shapes.Item(1).CanvasItems.Range(1)         rngCanvasShapes.Select()         rngCanvasShapes.Delete() }` |

**成员属性**

#### **GroupShapes.Application**

返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **GroupShapes** 对象的变量。

#### **GroupShapes.Count**

返回一个 **Long** 类型的值，该值代表集合中图形的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **GroupShapes** 对象的变量。

#### **GroupShapes.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **GroupShapes** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **GroupShapes.Parent**

返回一个 **Object** 类型值，该值代表指定 **GroupShapes** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **GroupShapes** 对象的变量。

适用环境：web

适用平台：windows/linux