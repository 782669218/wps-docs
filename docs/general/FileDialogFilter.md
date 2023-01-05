#### **FileDialogFilter**



返回通过 **FileDialog** 对象显示的文件对话框中的一个文件筛选。每个文件筛选确定文件对话框中显示的文件。

**说明**

使用 **FileDialogFilters** 集合的 **Item** 方法可返回一个 **FileDialogFilter** 对象。使用 **Add** 方法可向 **FileDialogFilters** 集合中添加一个 **FileDialogFilter** 对象。使用 **Extensions** 属性可返回 **FileDialogFilter** 对象用于筛选文件的扩展名，使用 **Description** 属性可以返回筛选的说明，但是这两个属性均为只读。如果要设置扩展名和说明，必须使用 **Add** 方法。

**示例**

下面的示例循环访问**“SaveAs”**对话框的默认筛选器，并显示每个包括 ET 文件的筛选器的说明。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {      //Declare a variable as a FileDialogFilters collection.      //Declare a variable as a FileDialogFilter object.      //Set the FileDialogFilters collection variable to     //the FileDialogFilters collection of the SaveAs dialog box.     let fdfs = Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters      //Iterate through the description and extensions of each     //default filter in the SaveAs dialog box.     for (let i = 1; i <= fdfs.Count; i++) {         let fdf = fdfs.Item(i)         //Display the description of filters that include         //ET files.         if (fdf.Extensions.indexOf("xls") > -1) {             alert("Description of filter: " + fdf.Description)         }     } }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 Application 对象，代表 **FileDialogFilter** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **FileDialogFilter** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Description** | 获取一个 **String** 类型的值，代表每个 **Filter** 对象的说明。该说明为文件对话框中显示的文本。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Extensions**  | 获取一个包含扩展名的值，这些扩展名决定对于每个 **Filter** 对象在文件对话框中显示的文件。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **FileDialogFilter** 对象的 **Parent** 对象。只读。     |

**成员属性**

#### **FileDialogFilter.Application**

获取一个 Application 对象，代表 **FileDialogFilter** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **FileDialogFilter** 对象的变量。

**说明**

#### **FileDialogFilter.Creator**

获取一个 32 位整数，指示创建 **FileDialogFilter** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **FileDialogFilter** 对象的变量。

#### **FileDialogFilter.Description**

获取一个 **String** 类型的值，代表每个 **Filter** 对象的说明。该说明为文件对话框中显示的文本。只读。

**语法**

**express.Description**

*express*   一个代表 **FileDialogFilter** 对象的变量。

**示例**

**示例**

下面的示例循环访问**“SaveAs”**对话框的默认筛选器，并显示每个包括 ET 文件的筛选器的说明。Extensions 属性用于查找相应的筛选对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {      //Declare a variable as a FileDialogFilters collection.      //Declare a variable as a FileDialogFilter object.      //Set the FileDialogFilters collection variable to     //the FileDialogFilters collection of the SaveAs dialog box.     let fdfs = Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters      //Iterate through the description and extensions of each     //default filter in the SaveAs dialog box.     for(let i = 1;i <= fdfs.Count;i++) {         let fdf = fdfs.Item(i)         //Display the description of filters that include         //ET files         if(fdf.Extensions.indexOf("xls") > -1) {             alert("Description of filter: " + fdf.Description)         }     }  }` |

#### **FileDialogFilter.Extensions**

获取一个包含扩展名的值，这些扩展名决定对于每个 **Filter** 对象在文件对话框中显示的文件。只读。

**语法**

**express.Extensions**

*express*   一个代表 **FileDialogFilter** 对象的变量。

**示例**

下面的示例通过循环访问**“SaveAs”**对话框中的筛选器来显示 ET 文件的扩展名和说明。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {      //Declare a variable as a FileDialogFilters collection.      //Declare a variable as a FileDialogFilter object.      //Set the FileDialogFilters collection variable to     //the FileDialogFilters collection of the SaveAs dialog box.     let fdfs = Application.FileDialog(Application.Enum.msoFileDialogSaveAs).Filters      //Iterate through the description and extensions of each     //default filter in the SaveAs dialog box.     for(let i = 1;i <= fdfs.Count;i++) {         let fdf = fdfs.Item(i)         //Display the description of filters that include         //ET files         if(fdf.Extensions.indexOf("xls") > -1) {             alert("Description of filter: " + fdf.Description)         }     }  }` |

#### **FileDialogFilter.Parent**

获取 **FileDialogFilter** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **FileDialogFilter** 对象的变量。

适用环境：web

适用平台：windows/linux