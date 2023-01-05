**FileExportConverter**



代表用于保存文件的文件转换器。

**说明**

无法新建文件转换器，也无法向 **FileExportConverters** 集合中添加文件转换器。**FileExportConverter** 对象是在安装 WPS Office 时或安装补充文件转换器时添加的。

使用 **FileExportConverters**(*Index*) 可返回单个 **FileExportConverter** 对象，其中 *Index* 为整数。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示与该集合中第二个 ET 工作表转换器关联的扩展名。*/ alert(FileExportConverters.Item(2).Extensions)` |

索引号代表文件转换器在 **FileExportConverters** 集合中的位置。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示该集合中第一个文件转换器的说明。*/ alert(FileExportConvters.Item(1).Description)` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回代表 ET 应用程序的 **Application** 对象。只读。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回表示在其中创建特定对象的应用程序的 32 位整数。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Description** | 返回文件转换器的说明。**String** 类型，只读                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Extensions**  | 返回与指定 **FileExportConverter** 对象关联的文件扩展名。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FileFormat**  | 返回一个整数，该整数标识与指定 **FileExportConverter** 对象关联的文件格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型的值，该值代表指定 **FileExportConverter** 对象的父对象。只读。 |

**成员属性**

#### **FileExportConverter.Application**

返回代表 ET 应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **FileExportConverter** 对象的变量。

#### **FileExportConverter.Creator**

返回表示在其中创建特定对象的应用程序的 32 位整数。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **FileExportConverter** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回十六进制值 5843454C，该值表示字符串“XCEL”。**Creator** 属性设计为在 ET for the Macintosh 中使用，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **FileExportConverter.Description**

返回文件转换器的说明。**String** 类型，只读

**语法**

**express.Description**

*express*   一个代表 **FileExportConverter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例显示 FileExportConverters 集合中第一个文件转换器的说明。*/ function test(){ let fcTemp = FileExportConverter.Item(1) alert(fcTemp.Description) }` |

#### **FileExportConverter.Extensions**

返回与指定 **FileExportConverter** 对象关联的文件扩展名。**String** 类型，只读。

**语法**

**express.Extensions**

*express*   一个代表 **FileExportConverter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例显示 FileExportConverters 集合中第一个文件转换器的文件扩展名。*/ function test(){ let fcTemp = FileExportConverters.Item(1) alert("The file name extensions for the file converter are: " + fcTemp.Extensions) }` |

#### **FileExportConverter.FileFormat**

返回一个整数，该整数标识与指定 **FileExportConverter** 对象关联的文件格式。只读。

**语法**

**express.FileFormat**

*express*   一个代表 **FileExportConverter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例显示 FileExportConverters 集合中第一个文件转换器的文件格式标识符。*/ function test(){ let fcTemp = FileExportConverters.Item(1) alert("The file format identifier for the file converter is: " + fcTemp.FileFormat) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例说明在使用 FileExportConverters 集合中的第一个文件转换器保存文件时，如何将文件格式标识符用作 Workbook 对象的 SaveAs 方法中的参数。*/ Application.ActiveWorkbook.SaveAs("C:\\temp\\myFile.xyz", Application.FileExportConverters.Item(1).FileFormat, null, null, null, false)  ` |

#### **FileExportConverter.Parent**

返回一个 **Object** 类型的值，该值代表指定 **FileExportConverter** 对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **FileExportConverter** 对象的变量。

适用环境：web

适用平台：windows/linux