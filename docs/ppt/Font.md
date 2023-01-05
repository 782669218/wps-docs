**Font**



包含对象的字体属性（如字体名称、字号、颜色等）。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllCaps**     | 如果字体格式为全部字母大写，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Bold**        | 如果字体格式为加粗，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Color**       |                                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Italic**      | 如果将字体或范围设置为倾斜格式，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回或设置指定对象的名称。**String** 类型，可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**      | 如果将指定字体设置为阴影格式，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Size**        | 返回或设置字号（以磅为单位）。**Single** 类型，可读写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Subscript**   | 如果字体格式为下标，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Superscript** | 如果字体格式为上标，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Underline**   | 返回或设置应用于字体的下划线类型。可读写 **WdUnderline** 类型。 |

**成员属性**

#### **Font.AllCaps**

如果字体格式为全部字母大写，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.AllCaps**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdU****ndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。该属性可设置为 **True**、**False** 或 **wdToggle**（与当前设置相反）。

如果设置 **AllCaps** 为 **True**，则 **SmallCaps** 属性的值为 **False**，反之亦然。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动文档第三段的文本格式是否设置为全部字母大写。*/ function test() {   if(Application.ActiveDocument.Paragraphs.Item(3).Range.Font.AllCaps == true){        alert("Text is all caps.")   }   else{       alert("Text is not all caps.")   } }  /*本示例将选定文本的字体格式设置为全部字母大写。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal){       Application.Selection.Font.AllCaps = true   }   else{       alert("You need to select some text.")   } }` |

#### **Font.Bold**

如果字体格式为加粗，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Bold**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：如果选定内容中有一部分字体为加粗格式，则将全部选定内容的格式都设为加粗格式。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal) {       if(Application.Selection.Font.Bold == wdUndefined) {           Application.Selection.Font.Bold = true       }   }   else{       alert("You need to select some text.")   } }` |

#### **Font.Color**

**语法**

**express.Color**

*express*   一个代表 **Font** 对象的变量。

#### **Font.Italic**

如果将字体或范围设置为倾斜格式，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Italic**

*express*   一个代表 **Font** 对象的变量。

**说明**

该属性返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例检查所选内容中的倾斜格式并删除所选内容中存在的倾斜格式。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal){       let mySel = Application.Selection.Font.Italic        if(mySel == wdUndefined || mySel == true){           alert("there is italic text in selection. " + "Click OK to remove.")           Application.Selection.Font.Italic = false       }       else{           alert("No italic text in the selection.")       }   }   else{       alert("You need to select some text.")   } }` |

#### **Font.Name**

返回或设置指定对象的名称。**String** 类型，可读写。

**语法**

**express.Name**

*express*   一个代表 **Font** 对象的变量。

**说明**

本示例将选定内容设置为 Arial 字体的加粗格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例将选定内容设置为 Arial 字体的加粗格式。 function test() {   Application.Selection.Font.Name = "Arial"   Application.Selection.Font.Bold = true }` |

#### **Font.Shadow**

如果将指定字体设置为阴影格式，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.Shadow**

*express*   一个代表 **Font** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对所选内容应用阴影和加粗格式。/* function test() {   if(Application.Selection.Type == wdSelectionNormal){       Application.Selection.Font.Shadow = true       Application.Selection.Font.Bold = true   }   else {       alert("You need to select some text.")   } }` |

#### **Font.Size**

返回或设置字号（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.Size**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例插入文本并将该文本的第七个单词的字号设置为 20 磅。*/ function test() {   Application.Selection.Collapse(wdCollapseEnd)   let rng = Application.Selection.Range   rng.Font.Reset()   rng.InsertBefore( "This is a demonstration of font size.")   rng.Words.Item(7).Font.Size = 20 }  /*以下示例确定所选文本的字号。*/ function test() {   let mySel = Application.Selection.Font.Size   if(mySel == wdUndefined){       alert("there is a mix of font sizes in the selection.")   }   else {       alert(mySel + " points")   } }` |

#### **Font.Subscript**

如果字体格式为下标，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Subscript**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

如果将 **Subscript** 属性设为 **True**，则 **Superscript** 属性会设为 **False**，反之亦然。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的开头插入文本，并将第十个字符设为下标。*/ function test() {   let myRange = Application.ActiveDocument.Range(0, 0)   myRange.InsertAfter("Water = H20")   myRange.Characters.Item(10).Font.Subscript = true }  /*本示例检查所选文本的下标格式。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal){       let mySel = Application.Selection.Font.Subscript       if(mySel == wdUndefined || mySel == true){           alert("Subscript text exists in the selection.")       }       else{           alert("No subscript text in the selection.")       }   }   else{       alert("You need to select some text.")   } }` |

#### **Font.Superscript**

如果字体格式为上标，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Superscript**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

如果将 **Superscript** 属性设为 **True**，则 **Subscript** 属性会设为 **False**，反之亦然。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的开头插入文本，并将第四个单词的两个字符设为上标。*/ function test() {   let myRange = Application.ActiveDocument.Range(0, 0)   myRange.InsertAfter("Superscript in the 4th word.")   Application.ActiveDocument.Range(20, 22).Font.Superscript = true }  /*本示例将所选文本设为上标。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal){       Application.Selection.Font.Superscript = true   }   else{       alert("You need to select some text.")   } }` |

#### **Font.Underline**

返回或设置应用于字体的下划线类型。可读写 **WdUnderline** 类型。

**语法**

**express.Underline**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例给所选内容添加单下划线。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal){       Application.Selection.Font.Underline = wdUnderlineSingle   }   else{       alert("You need to select some text.")   } }` |

适用环境：web

适用平台：windows/linux