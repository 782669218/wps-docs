**Slides**



指定的演示文稿中所有 [**Slide**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Slide/Slide%20.htm#jsObject_Slide) 对象的集合。

**说明**

以下示例说明如何执行下列操作：

- 创建一张幻灯片并将其添加到集合中
- 返回通过名称、索引号或幻灯片 ID 号指定的单张幻灯片
- 返回演示文稿中幻灯片的子集
- 同时将一种属性或方法应用到演示文稿中的所有幻灯片

## 示例

使用 [**Slides**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Slides/Slides%20.htm#jsObject_Slides) 属性返回 **Slides** 集合。使用 [**Add**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Slides/Slides%20.htm#Slides.Add) 方法新建幻灯片并将其添加到集合中。以下示例将新幻灯片添加到活动演示文稿中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Add(2, ppLayoutBlank)` |

使用 

Slides

(

index

)（其中 

index

 是幻灯片名称或索引号）或使用 

Slides.FindBySlideID

(

index

)（其中 

index

 是幻灯片 ID 号）返回单个 

Slide

 对象。以下示例设置活动演示文稿中第一张幻灯片的版式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Layout = ppLayoutTitle` |

使用 **Slides.Range**(*index*) 返回代表 **Slides** 集合的一个子集的 [**SlideRange**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/SlideRange/SlideRange%20.htm#jsObject_SlideRange) 对象，其中 *index* 是幻灯片索引号或名称，或者幻灯片索引号的数组或幻灯片名称的数组。

如果要同时对演示文稿中所有幻灯片进行某种操作（例如全部删除或设置它们的某些属性），可不带参数使用 **Slides.Range** 创建一个包含 **Slides** 集合中所有幻灯片的 **SlideRange** 集合，然后将适当的属性或方法应用于 **SlideRange** 集合。

**方法**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**            | 创建新的幻灯片，将其添加到 Slides 集合，然后返回幻灯片。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddSlide**       | 创建新的幻灯片，将其添加到 Slides 集合，然后返回幻灯片。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FindBySlideID**  | 返回一个 **Slide** 对象，该对象代表具有指定幻灯片 ID 号的幻灯片。每张幻灯片在创建时都会自动分配唯一的幻灯片 ID 号。使用 **SlideID** 属性可返回幻灯片的 ID 号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertFromFile** | 将来自某文件的幻灯片插入到演示文稿中的指定位置。返回一个代表所插入幻灯片数量的 **integer**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**           | 从指定集合中返回单个对象。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Paste**          | 将剪贴板上的幻灯片粘贴到演示文稿的 **Slides** 集合中。用 **Index** 参数可以指定要插入幻灯片的位置。返回一个代表粘贴对象的 **SlideRange** 对象。每张粘贴幻灯片都成为指定的 **Slides** 集合的成员。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Range**          | 返回一个代表 **Slides** 集合中的幻灯片子集的 **SlideRange** 对象。 |

**属性**

|                                                              | 名称            | 说明                                                        |
| ------------------------------------------------------------ | --------------- | ----------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                  |

**成员方法**

#### **Slides.Add**

创建新的幻灯片，将其添加到 Slides 集合，然后返回幻灯片。

**语法**

**express.Add(Index, PpSlideLayout)**

*express*   一个代表 **Slides** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**      | **说明**               |
| --------------- | ------------- | ----------------- | ---------------------- |
| *Index*         | 必选          | **Number**        | 要添加的幻灯片的索引。 |
| *PpSlideLayout* | 必选          | **PpSlideLayout** | 幻灯片的版式。         |

**说明**

| PpSlideLayout 可以是下列 PpSlideLayout 常量之一。 |
| ------------------------------------------------- |
| **ppLayoutBlank**                                 |
| **ppLayoutChart**                                 |
| **ppLayoutChartAndText**                          |
| **ppLayoutClipartAndText**                        |
| **ppLayoutClipArtAndVerticalText**                |
| **ppLayoutFourObjects**                           |
| **ppLayoutLargeObject**                           |
| **ppLayoutMediaClipAndText**                      |
| **ppLayoutMixed**                                 |
| **ppLayoutObject**                                |
| **ppLayoutObjectAndText**                         |
| **ppLayoutObjectOverText**                        |
| **ppLayoutOrgchart**                              |
| **ppLayoutTable**                                 |
| **ppLayoutText**                                  |
| **ppLayoutTextAndChart**                          |
| **ppLayoutTextAndClipart**                        |
| **ppLayoutTextAndMediaClip**                      |
| **ppLayoutTextAndObject**                         |
| **ppLayoutTextAndTwoObjects**                     |
| **ppLayoutTextOverObject**                        |
| **ppLayoutTitle**                                 |
| **ppLayoutTitleOnly**                             |
| **ppLayoutTwoColumnText**                         |
| **ppLayoutTwoObjectsAndText**                     |
| **ppLayoutTwoObjectsOverText**                    |
| **ppLayoutVerticalText**                          |
| **ppLayoutVerticalTitleAndText**                  |
| **ppLayoutVerticalTitleAndTextOverChart**         |

#### **Slides.AddSlide**

创建新的幻灯片，将其添加到 Slides 集合，然后返回幻灯片。

**语法**

**express.AddSlide(Index, pCustomLayout)**

*express*   一个代表 **Slides** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**     | **说明**               |
| --------------- | ------------- | ---------------- | ---------------------- |
| *Index*         | 必选          | **Integer**      | 要添加的幻灯片的索引。 |
| *pCustomLayout* | 必选          | **CustomLayout** | 幻灯片的版式。         |

#### **Slides.FindBySlideID**

返回一个 **Slide** 对象，该对象代表具有指定幻灯片 ID 号的幻灯片。每张幻灯片在创建时都会自动分配唯一的幻灯片 ID 号。使用 **SlideID** 属性可返回幻灯片的 ID 号。

**语法**

**express.FindBySlideID(SlideID)**

*express*   一个代表 **Slides** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                    |
| --------- | ------------- | ------------ | ----------------------------------------------------------- |
| *SlideID* | 必选          | **Number**   | 指定要返回的幻灯片的 ID 号。WPP在创建幻灯片时会分配该编号。 |

**说明**

与 **SlideIndex** 属性不同，在演示文稿中添加或重新排列幻灯片时，**Slide** 对象的 **SlideID** 属性不会改变。因此，与对 **Item** 方法使用幻灯片索引号相比，对 **FindBySlideID** 方法使用幻灯片 ID 号可以更可靠地从 **Slides** 集合返回特定的 **Slide** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例示范如何检索 Slide 对象的唯一 ID 号，并使用此编号从 Slides 集合返回该 Slide 对象。*/ function test() {     let gslides = Application.ActivePresentation.Slides      // Get slide ID     let graphSlideID = gslides.Add(2, ppLayoutChart).SlideID          // Use ID to return specific slide     gslides.FindBySlideID(graphSlideID).SlideShowTransition.EntryEffect = ppEffectCoverLeft }` |

#### **Slides.InsertFromFile**

将来自某文件的幻灯片插入到演示文稿中的指定位置。返回一个代表所插入幻灯片数量的 **integer**。

**语法**

**express.InsertFromFile(FileName, Index, SlideStart, SlideEnd)**

*express*   一个代表 **Slides** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*   | 必选          | **String**   | 指示包含要插入的幻灯片的文件的文件名。                       |
| *Index*      | 必选          | **Number**   | 要在其后插入新幻灯片的 Slide 对象在指定 Slides 集合中的索引号。 |
| *SlideStart* | 可选          | **Number**   | 第一个 Slide 对象在 Slides 集合中的索引号，该集合在由 FileName 指定的文件中。 |
| *SlideEnd*   | 可选          | **Number**   | 最后一个 Slide 对象在 Slides 集合中的索引号，该集合在由 FileName 指定的文件中。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动演示文稿的第二张灯片后插入文件 C:\Ppt\Sales.ppt 的第三张到第六张幻灯片。*/ Application.ActivePresentation.Slides.InsertFromFile("C:\\ppt\\sales.ppt", 2, 3, 6)` |

#### **Slides.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Slides** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Variant**  | 集合中要返回的单个对象的名称或索引号。 |

#### **Slides.Paste**

将剪贴板上的幻灯片粘贴到演示文稿的 **Slides** 集合中。用 **Index** 参数可以指定要插入幻灯片的位置。返回一个代表粘贴对象的 **SlideRange** 对象。每张粘贴幻灯片都成为指定的 **Slides** 集合的成员。

**语法**

**express.Paste(Index)**

*express*   一个代表 **Slides** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 可选          | **Number**   | 表示要将剪贴板上的幻灯片粘贴在它前面的幻灯片的索引号。如果省略此参数，则剪贴板上的幻灯片将粘贴到演示文稿的最后一张幻灯片之后。 |

**说明**

将剪贴板内容粘贴到窗口中之前，请使用 **ViewType** 属性设置窗口视图。下表显示了可以粘贴到每个视图中的内容。

| 视图                   | 可以从剪贴板中粘贴以下内容                                   |
| ---------------------- | ------------------------------------------------------------ |
| 幻灯片视图或备注页视图 | 形状、文本或整张幻灯片。如果从剪贴板粘贴一个幻灯片，该幻灯片的图像将作为嵌入对象被插入到幻灯片、母版或备注页中；如果选中形状，粘贴的文本将附加到形状文本之后；如果选中文本，粘贴的文本将替换选中的文本；如果选中任何其他对象，粘贴的文本将被放到它自己的文本框中。粘贴的形状将被放到 z 顺序的最上端且不会替换选中的形状。 |
| 大纲视图               | 文本或整张幻灯片。不能向大纲视图粘贴形状。粘贴的幻灯片将被插到光标所在的幻灯片之前。 |
| 幻灯片浏览视图         | 整张幻灯片。不能向幻灯片浏览视图粘贴形状或文本。粘贴的幻灯片将被插到光标处或演示文稿中最后选中的一张幻灯片之后。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例从演示文稿“Old Sales”中剪切第三张和第五张幻灯片，然后将它们插入到活动演示文稿中的第四张幻灯片之前。*/ function test() {     Application.Presentations.Item("Old Sales").Slides.Range([3, 5]).Cut()     Application.ActivePresentation.Slides.Paste(4) }` |

#### **Slides.Range**

返回一个代表 **Slides** 集合中的幻灯片子集的 **SlideRange** 对象。

**语法**

**express.Range(Index)**

*express*   一个代表 **Slides** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 可选          | **Variant**  | 要包含在范围内的各个幻灯片。可以是 Integer 类型的值（指定幻灯片的索引号）、String 类型的值（指定幻灯片的名称），或者是包含整数或字符串的数组。如果省略此参数，则 Range 方法将返回指定集合中的所有对象。 |

**说明**

虽然可以使用 **Range** 方法返回任意数目的形状或幻灯片，但如果只需要返回该集合的一个成员，则使用 **Item** 方法会更为简单。例如，Shapes(1) 比 Shapes.Range(1) 简单，Slides(2) 比 Slides.Range(2) 简单。

**成员属性**

#### **Slides.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Slides** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某函数。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(pptPres)  {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")   }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for(let i = 1; i <= shpOle.Count; i++)      {         if(shpOle.Item(i).Type == msoLinkedOLEObject)          {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **Slides.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **Slides** 对象的变量。

适用环境：web

适用平台：windows/linux