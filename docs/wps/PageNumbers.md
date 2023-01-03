#### **PageNumbers**



**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 返回一个 **PageNumber** 对象，该对象代表添加到节中的页眉或页脚中的页码。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **PageNumber** 对象。                       |

**属性**

|                                                              | 名称      | 说明                                                       |
| ------------------------------------------------------------ | --------- | ---------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count** | 返回一个 **Long** 类型的值，该值代表集合中的页码数。只读。 |

**成员方法**

#### **PageNumbers.Add**

返回一个 **PageNumber** 对象，该对象代表添加到节中的页眉或页脚中的页码。

**语法**

**express.Add(PageNumberAlignment, FirstPage)**

*express*   一个代表 **PageNumbers** 对象的变量。

**参数**

| **名称**              | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *PageNumberAlignment* | 可选          | **Variant**  | 可以是任意 WdPageNumberAlignment 常量。                      |
| *FirstPage*           | 可选          | **Variant**  | 如果该属性值为 False，则文档中的第一页的页眉和页脚与后面各页的页眉和页脚都不相同；如果将 FirstPage 设置为 False，则在第一页不添加页码；如果省略该参数，则由 DifferentFirstPageHeaderFooter 属性控制该设置。 |

**说明**

如果将 **HeaderFooter** 对象的 **LinkToPrevious** 属性设置为 **True**，则文档中各节的页码都将连续编号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档第一节的主页脚添加页码*/ function test() {  let Sec = Application.ActiveDocument.Sections(1)      Sec.Footers(wdHeaderFooterPrimary).PageNumbers.Add(wdAlignPageNumberLeft,true) }  /*本示例在活动文档的页眉中创建页码并进行格式设置。*/ function test() { 	let myPgNum = Application.ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.Add(wdAlignPageNumberCenter,true)     myPgNum.Select()     Selection.Range.Italic = true     Selection.Range.Bold = true }` |

#### **PageNumbers.Item**

返回集合中的单个 **PageNumber** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **PageNumbers** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

**成员属性**

#### **PageNumbers.Count**

返回一个 **Long** 类型的值，该值代表集合中的页码数。只读。

**语法**

**express.Count**

*express*   一个代表 **PageNumbers** 对象的变量。

适用环境：web

适用平台：windows/linux