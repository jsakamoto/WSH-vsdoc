var WScript = {
    Echo: function (text) { },
    Quit: function () { },
    ScriptFullName: ""
};

var ActiveXObject = function (clsid) {
    ///<param name="clsid">
    /// Specify CLSID of COM object to create. For example:
    ///<para>"Scripting.FileSystemObject"</para>
    ///<para>"WScript.Shell"</para>
    ///<para>"Scripting.Dictionary"</para>
    ///</param>
    if (clsid == "Scripting.FileSystemObject") {
        var TextStream = function () {
            this.AtEndOfStream = false;
            this.Close = function () { };
            this.ReadAll = function () { return ""; };
        };

        var File = function () {
            this.Name = "";
            this.Size = 0;

            this.OpenAsTextStream = function (iomode, format) {
                /// <param name="iomode" type="int">
                /// Optional. Can be one of three constants:
                /// <para>1 : ForReading - Open a file for reading only. You can't write to this file.</para>
                /// <para>2 : ForWriting - Open a file for writing.</para>
                /// <para>8 : ForAppending - Open a file and write to the end of the file.</para>
                /// </param>
                /// <param name="format" type="int">
                /// Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.Tristate values:
                /// <para>-2 : TristateUseDefault - Opens the file using the system default.</para>
                /// <para>-1 : TristateTrue - Opens the file as Unicode.</para>
                /// <para> 0 : TristateFalse - Opens the file as ASCII.</para>
                /// </param>
                /// <returns type="TextStream" />
                return new TextStream();
            }
        };

        var Folder = function () {
            this.Name = "";
            this.Size = 0;
            this.Files = [new File()];
            this.Move = function (destinationPath) { };
        };
        Folder.prototype.SubFolders = [new Folder()];

        this.GetFolder = function (fullPath) { return new Folder(); };
        this.FileExists = function (fullPath) {
            /// <returns type="bool" />
            return true;
        };
        this.CopyFile = function (sourcePath, destinationPath, overwrite) {
        	/// <param name="sourcePath" type="String"></param>
        	/// <param name="destinationPath" type="String"></param>
        	/// <param name="overwrite" optional="true" type="bool"></param>
        };

        this.FolderExists = function (fullPath) { return true; };
        this.MoveFolder = function (sourceFullPath, destinationFullPath) { };
        this.CreateFolder = function (fullPath) { };
        this.GetSpecialFolder = function (folderspec) {
            /// <param name="folderspec" type="int">
            /// Required. The name of the special folder to be returned. Can be any of the constants:
            /// <para>0 : WindowsFolder - The Windows folder contains files installed by the Windows operating system.</para>
            /// <para>1 : SystemFolder - The System folder contains libraries, fonts, and device drivers.</para>
            /// <para>2 : TemporaryFolder - The Temp folder is used to store temporary files. Its path is found in the TMP environment variable.</para>
            /// </param>
        };
        this.BuildPath = function (path, name) {
            /// <param name="path" type="String">Required. Existing path to which name is appended. Path can be absolute or relative and need not specify an existing folder.</param>
            /// <param name="name" type="String">Required. Name being appended to the existing path.</param>
            /// <returns type="String" />
            return "";
        };
        this.GetFile = function (filespec) {
            /// <param name="filespec" type="String">Required. The filespec is the path (absolute or relative) to a specific file.</param>
            /// <returns type="File" />
            return new File();
        };
        this.OpenTextFile = function (filename, iomode, create, format) {
            /// <param name="filename" type="String">Required. String expression that identifies the file to open.</param>
            /// <param name="iomode" type="int">
            /// Optional. Can be one of three constants:
            /// <para>1 : ForReading - Open a file for reading only. You can't write to this file.</para>
            /// <para>2 : ForWriting - Open a file for writing.</para>
            /// <para>8 : ForAppending - Open a file and write to the end of the file.</para>
            /// </param>
            /// <param name="create" type="bool">Optional. Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created.</param>
            /// <param name="format" type="int">
            /// Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.Tristate values:
            /// <para>-2 : TristateUseDefault - Opens the file using the system default.</para>
            /// <para>-1 : TristateTrue - Opens the file as Unicode.</para>
            /// <para> 0 : TristateFalse - Opens the file as ASCII.</para>
            /// </param>
            /// <returns type="TextStream" />
            return new TextStream();
        };
    }

    if (clsid == "WScript.Shell") {
        this.Run = function (command, windowStyle, waitOnReturn) {
            ///<param name="command" type="String"></param>
            ///<param name="windowStyle" type="int" optional="true">0:hidden, 1:restore, 2:minimize, 3:maximize, ...</param>
            ///<param name="waitOnReturn" type="bool" optional="true"></param>
        }

        // for Registry
        // http://msdn.microsoft.com/en-us/library/x05fawxd.aspx
        this.RegRead = function (strName) {
            /// <param name="strName" type="String">String value indicating the key or value-name whose value you want.</param>
        }
        // http://msdn.microsoft.com/en-us/library/yfdfhz1b.aspx
        this.RegWrite = function (strName, anyValue, strType) {
            /// <param name="strName" type="String">String value indicating the key-name, value-name, or value you want to create, add, or change.</param>
            /// <param name="anyValue" type="Object">The name of the new key you want to create, the name of the value you want to add to an existing key, or the new value you want to assign to an existing value-name.</param>
            /// <param name="strType" type="String">Optional. String value indicating the value's data type, REG_SZ, REG_DWORD, REG_BINARY, REG_EXPAND_SZ. The REG_MULTI_SZ type is not supported for the RegWrite method.</param>
        }
    }

    if (clsid == "Scripting.Dictionary")
    {
        // http://msdn.microsoft.com/en-us/library/x4k5wbx4.aspx
        this.Add = function (key, item) { };
        this.Items = function () { return new VBArray(null); };
    }
};

var Enumerator = function (enumerable) {
    this.atEnd = function () { return false; };
    this.moveNext = function () { };
    this.item = function () { return enumerable[0]; };
};

var GetObject = function (pathname) {
    /// <summary>
    /// </summary>
    /// <param name="pathname" type="String">ex) 'IIS://localhost/W3SVC/1/Root'</param>

    var getObject = function (pathname, param) {
    	/// <summary></summary>
        /// <param name="pathname">ex) 'IIsWebVirtualDir'</param>
    	/// <param name="param">ex) url</param>
    	/// <returns type=""></returns>
        if (pathname.toLowerCase() == 'iiswebvirtualdir') {
            // http://msdn.microsoft.com/en-us/library/ms524579.aspx
            return {
                SetInfo: function () { },
                Path: '',
                AppPoolId: '',
                EnableDirBrowsing: true,
                AccessRead: true,
                AccessWrite: false,
                AccessSource: false,
                AccessScript: false,
                AccessExecute: false,
                DontLog: false,
                AuthAnonymous: true,
                ///<summary>Oh</summary>
                AuthBasic: true,
                AuthNTLM: false,
                AppCreate2: function (type) {
                	/// <summary></summary>
                    /// <param name="type" type="Number">
                    /// <para>0: In Proc</para>
                    /// <para>1: Out Of Proc</para>
                    /// <para>2: Pooled Out Of Proc</para>
                    /// </param>
                },
                AnonymousPasswordSync: true,
                AnonymousUserName: '',
                AnonymousUserPass: '',
                EnableDefaultDoc: true,
                DefaultDoc: '',
                AdsPath: '',
                ScriptMaps: new VBArray(null)
            };
        }
        return {};
    };

    if (pathname.toLowerCase().substr(0, 4) == 'iis:') {
        return {
            GetObject: getObject,
            Create: getObject
        };
    }

    if (pathname.toLowerCase().substr(0, 6) == 'winnt:') {
        return {
            Create: function (pathname, param) {
                return {
                    SetInfo: function () { },
                    SetPassword: function (passwordText) { },
                    userFlags: 0
                };
            },
            Add: function (pathname) { }
        };
    }

    return {};
};