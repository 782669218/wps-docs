**Fonts**



指定的演示文稿中所有 **Font**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font)对象的集合。

**说明**

每个 [**Font** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font)对象代表演示文稿中使用的一种字体。

“Genigraphics 向导”使用 **Fonts** 集合来决定，当 Genigraphics 转换幻灯片时是否在指定演示文稿中有任何字体不受支持。如果只需要设置特定项目符号或文本范围的字符格式，请使用 [**Font** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font)属性为项目符号或文本范围返回 [**Font** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font)对象。

“Genigraphics 向导”使用户可直接将演示文稿传输到 Genigraphics，以便转换成胶片幻灯片、投影机透明胶片或其他专用媒体格式。有关 Genigraphics 所提供的服务的详细信息，请访问 Genigraphics 网站：http://www.genigraphics.com/。此服务仅在美国国内可用。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用 Fonts 属性返回 Fonts 集合。以下示例显示活动演示文稿中使用的字体种数。*/ alert(Application.ActivePresentation.Fonts.Count)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用 Fonts(index) 返回单个 Font 对象，其中 index 是字体名称或索引号。以下示例检查活动演示文稿的第一种字体是否已嵌入。*/ function test() {     if(Application.ActivePresentation.Fonts.Item(1).Embedded == true) {         alert("Font 1 is embedded")     } }` |

**方法**

|                                                              | 名称        | 说明                                       |
| ------------------------------------------------------------ | ----------- | ------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**    | 从指定集合中返回单个对象。返回值类型为Font |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Replace** | 替换 **Fonts** 集合中的一种字体。          |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                   |

**成员方法**

#### **Fonts.Item**

从指定集合中返回单个对象。返回值类型为Font

**语法**

**express.Item(Index)**

*express*   一个代表 **Fonts** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Variant**  | 集合中要返回的单个对象的名称或索引号。 |

#### **Fonts.Replace**

替换 **Fonts** 集合中的一种字体。

**语法**

**express.Replace(Original, Replacement)**

*express*   一个代表 **Fonts** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**             |
| ------------- | ------------- | ------------ | -------------------- |
| *Original*    | 必选          | **String**   | 要替换的字体的名称。 |
| *Replacement* | 必选          | **String**   | 替换字体的名称。     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿中的 Times New Roman 字体替换为 Courier 字体。*/ function test() {     Application.ActivePresentation.Fonts.Replace("Times New Roman"，"Courier") }` |

**成员属性**

#### **Fonts.Application**

返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Fonts** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(pptPre) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {   let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes   for(let i =1;i<= Application.ActivePresentation.Slides.Item(1).Shapes.Count;i++) {       if(shpOle.Item(i).Type == msoLinkedOLEObject) {           alert(shpOle.Item(i).OLEFormat.Application.Name)       }									   } }` |

#### **Fonts.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **Fonts** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除窗口1外的所有窗口。*/ function test() {   for(let i=2; i <= Application.Windows.Count; i++){       Application.Windows.Item(i).Close()   } }` |

适用环境：web

适用平台：windows/linux