#### **EncryptionProvider**



提供用于设置权限、应用基础加密和解密的加密技术以及进行用户身份验证的方法。

**说明**

加密提供程序通过自定义的 COM 加载项实现。Office 文档内提供了存储区来存储加载项特定的信息，可以在其中存储加密、解密、应用权限和显示权限设置或验证用户界面所需的任何信息。

**方法**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Authenticate**      | 用于确定用户是否具有打开加密文档的适当权限。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CloneSession**      | 为即将保存的文件创建 **EncryptionProvider** 对象加密会话的另一个工作副本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DecryptStream**     | 对文档的加密数据流进行解密并将其返回。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **EncryptStream**     | 加密并返回文档的数据流。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **EndSession**        | 结束当前加密会话。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetProviderDetail** | 显示有关当前文档的加密的信息。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NewSession**        | 由 **EncryptionProvider** 对象在创建新的加密会话时使用。当文档被调入内存中后，提供程序将使用此会话缓存与加密、用户和权限相关的文档特定的信息。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Save**              | 保存加密的文档。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowSettings**      | 用于显示当前文档的加密设置的对话框。                         |

**成员方法**

#### **EncryptionProvider.Authenticate**

用于确定用户是否具有打开加密文档的适当权限。

**语法**

**express.Authenticate(ParentWindow, EncryptionData, PermissionsMask)**

*express*   一个代表 **EncryptionProvider** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型**   | **说明**                           |
| ----------------- | ------------- | -------------- | ---------------------------------- |
| *ParentWindow*    | 必选          | **IUnknown**   | 指定为了显示加密设置而调用的窗口。 |
| *EncryptionData*  | 必选          | **IUnknown**   | 包含当前文档的加密数据。           |
| *PermissionsMask* | 必选          | **无符号整数** | 加密提供程序加载项显示的用户界面。 |

**返回值**

Long

**说明**

它是 COM 加载项加密提供程序的显示位置，与适于应用加密的用户界面无关。例如，密码加密提供程序将提示用户输入密码。

#### **EncryptionProvider.CloneSession**

为即将保存的文件创建 **EncryptionProvider** 对象加密会话的另一个工作副本。

**语法**

**express.CloneSession(SessionHandle)**

*express*   一个代表 **EncryptionProvider** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**          |
| --------------- | ------------- | ------------ | ----------------- |
| *SessionHandle* | 必选          | **Long**     | 克隆的会话的 ID。 |

**说明**

此方法的输出是另一个具有相同加密设置的会话句柄。当执行自动保存操作时，将向您提供此会话句柄。

#### **EncryptionProvider.DecryptStream**

对文档的加密数据流进行解密并将其返回。

**语法**

**express.DecryptStream(SessionHandle, StreamName, EncryptedStream, UnencryptedStream)**

*express*   一个代表 **EncryptionProvider** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**           |
| ------------------- | ------------- | ------------ | ------------------ |
| *SessionHandle*     | 必选          | **Long**     | 当前会话的 ID。    |
| *StreamName*        | 必选          | **String**   | 数据流的 ID。      |
| *EncryptedStream*   | 必选          | **IUnknown** | 已加密数据流。     |
| *UnencryptedStream* | 必选          | **IUnknown** | 解密之前的数据流。 |

**说明**

此方法由 COM 加载项调用，调用时间为文档已打开并且在加载项已验证打开文档的用户已经过身份验证后。此方法是 EncryptStream 方法的逆向操作，它将加密数据转换回纯（未加密）数据。

#### **EncryptionProvider.EncryptStream**

加密并返回文档的数据流。

**语法**

**express.EncryptStream(SessionHandle, StreamName, UnencryptedStream, EncryptedStream)**

*express*   一个代表 **EncryptionProvider** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**                 |
| ------------------- | ------------- | ------------ | ------------------------ |
| *SessionHandle*     | 必选          | **Long**     | 当前会话的 ID。          |
| *StreamName*        | 必选          | **String**   | 加密的文档数据流的名称。 |
| *UnencryptedStream* | 必选          | **IUnknown** | 加密之前的数据流。       |
| *EncryptedStream*   | 必选          | **IUnknown** | 加密后的数据流信息。     |

**说明**

此方法通常是在执行保存操作的过程中由 COM 加载项调用。

#### **EncryptionProvider.EndSession**

结束当前加密会话。

**语法**

**express.EndSession(SessionHandle)**

*express*   一个代表 **EncryptionProvider** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**        |
| --------------- | ------------- | ------------ | --------------- |
| *SessionHandle* | 必选          | **Long**     | 当前会话的 ID。 |

**说明**

在执行保存操作期间，COM 加载项将调用 **CloneSession** 方法，以便为即将保存的文件创建 **EncryptionProvider** 对象加密会话的另一个工作副本。接下来调用 **Save** 方法，以获取要保留的有关加密设置的所有自定义信息。在以后再次打开此文档时，便可以使用此信息。然后调用 **EncryptStream** 方法，向提供程序提供文档的全部内容。最后，为克隆的会话句柄调用 **EndSession** 方法来完成该过程。

#### **EncryptionProvider.GetProviderDetail**

显示有关当前文档的加密的信息。

**语法**

**express.GetProviderDetail(encprovdet)**

*express*   一个代表 **EncryptionProvider** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**                 | **说明**             |
| ------------ | ------------- | ---------------------------- | -------------------- |
| *encprovdet* | 必选          | **EncryptionProviderDetail** | 指定所需的加密信息。 |

**返回值**

Variant

**说明**

通过此方法，可以查询 **EncryptionProvider** 对象以获取类似下面的信息：未安装您的自定义 COM 加载项的用户的下载 URL、实现的算法以及使用的加密模式。

#### **EncryptionProvider.NewSession**

由 **EncryptionProvider** 对象在创建新的加密会话时使用。当文档被调入内存中后，提供程序将使用此会话缓存与加密、用户和权限相关的文档特定的信息。

**语法**

**express.NewSession(ParentWindow)**

*express*   一个代表 **EncryptionProvider** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                           |
| -------------- | ------------- | ------------ | ---------------------------------- |
| *ParentWindow* | 必选          | **IUnknown** | 指定为了显示加密设置而调用的窗口。 |

**返回值**

Long

**说明**

此方法由 COM 加载项调用。

#### **EncryptionProvider.Save**

保存加密的文档。

**语法**

**express.Save(SessionHandle, EncryptionData)**

*express*   一个代表 **EncryptionProvider** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**        |
| ---------------- | ------------- | ------------ | --------------- |
| *SessionHandle*  | 必选          | **Long**     | 当前会话的 ID。 |
| *EncryptionData* | 必选          | **IUnknown** | 包含加密信息。  |

**返回值**

Long

**说明**

将文件保存为 Office Open XML 文件格式（它是支持自定义文件加密的唯一格式）时，COM 加载项将调用提供程序来加密文档。如果您尝试保存的目标文件格式不支持自定义文件加密，而您具有执行此操作的相应权限，则 WPS Office将在不加密的情况下保存该文档。这样便可以将文档导出为不支持加密或权限管理的格式。

#### **EncryptionProvider.ShowSettings**

用于显示当前文档的加密设置的对话框。

**语法**

**express.ShowSettings(SessionHandle, ParentWindow, ReadOnly, Remove)**

*express*   一个代表 **EncryptionProvider** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SessionHandle* | 必选          | **Long**     | 当前会话的 ID。                                              |
| *ParentWindow*  | 必选          | **IUnknown** | 指定为了显示加密设置而调用的窗口。                           |
| *ReadOnly*      | 必选          | **Boolean**  | 指定是否希望用户能更改加密设置。                             |
| *Remove*        | 必选          | **Boolean**  | 如果为 True，则在执行下一个保存操作的过程中将删除文档的加密。 |

**说明**

只能对已加密的文档调用此方法。您可以在 COM 加载项中使用此方法，以基于用户的任务显示您所希望的用户体验。例如，在纯加密情况中，可以显示一个对话框来更改文档的密码。在权限管理情况中，可以决定是显示用于更改权限的对话框，还是显示用户的权限。

适用环境：web

适用平台：windows/linux