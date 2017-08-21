//********************************************************************
//*
//*
//* Module Name:    Poster.js
//*
//* Abstract:       A assisant helps creating movie infomation files, such as *.nfo, poster.jpg.
//*					Poster.js Type FolderPath FileType
//*					    Type = txt, create *.nfo from *.txt file
//*                     Type = mov, download from douban
//*                     Type = clean, clean the file of *.nfo and *.jpg
//*                     Type = new, rename the folder by date which is created
//*
//********************************************************************
// Global declaration
                                                          
//----------------------------------------------------------------
// Start of localization Content
//----------------------------------------------------------------

// variable use to  concatenate  the Localization Strings.
// Error Messages
var L_InfoUnableToInclude_ErrorMessage      = "Error: Load module failed ";
var L_DeployTypeNotValid_ErrorMessage       = "Error: Invalid deploy type.";
var L_ProjectPathNotValid_ErrorMessage		= "Error: The project path is not valid.";

//-------------------------------------------------------------------------
// END of localization content
//-------------------------------------------------------------------------
// Define Version
var CONST_Program_Version		= "2015-12-25 10:09:04";

// Define constants
var CONST_ForReading            = 1;
var CONST_ForWriting            = 2;
var CONST_ForAppending          = 8;

var CONST_YES				  	= 1;
var CONST_NO				  	= 2;
var CONST_QUIT				    = 3;
var CONST_ALL				    = 4;

// Define the Exit Values
var EXIT_SUCCESS                = 0;
var EXIT_UNEXPECTED             = 255;
var EXIT_INVALID_INPUT          = 254;
var EXIT_METHOD_FAIL            = 250;
var EXIT_INVALID_PARAM          = 999;

// Define environment variables in DOS
var CONST_CHARS_IN_LINE 		= 78;
var CONST_SEPARATE_LINE_1		= "  ----------------------------------------------------------------------------";
var CONST_SEPARATE_LINE_2		= "    --------------------------------------------------------------------------";
var CONST_SEPARATE_LINE_3		= "      ------------------------------------------------------------------------";
var CONST_SEPARATE_LINE_4		= "        ----------------------------------------------------------------------";
var CONST_LIST_LEVEL_1			= "  ";
var CONST_LIST_LEVEL_2			= "     ";
var CONST_LIST_LEVEL_3			= "       ";
var CONST_LIST_LEVEL_4			= "         ";
var CONST_SEPARATE_SHORT        = "-----------------------------------------------------------------";

// Define execute result
var CONST_RESULT_OK 			= " [  OK  ]";
var CONST_RESULT_ERROR 			= " [ ERROR]";
var CONST_RESULT_FAILED 		= " [ FAIL ]";
var CONST_RESULT_IGNORE 		= " [IGNORE]";

// Define to display in table.
var CONST_SHOW_TABLE			= false;

// For blank line in  help usage
var EmptyLine_Text           	= " ";

// Supported movie type
var MOVIE_EXTENSION = "mkv|avi|rmvb|mp4|vob|ts|m2ts";

// Pause of Search in internet 
var PAUSE_DURATION = 1000;

// Constants for showing help for Usage
var L_ShowUsageLine01_Text      = " Scraper.js [-Version] ProjectPath DeployType";
var L_ShowUsageLine02_Text      = " Arguments:";
var L_ShowUsageLine03_Text      = "   Type      Type of scraper, such as mov, new";
var L_ShowUsageLine04_Text      = "   Path      Path of Folder.";
var L_ShowUsageLine05_Text      = "   NFO       1=single movie; 2=movie set; 3=TV Episodes.";

// Constant for *.nfo type
var CONST_SINGLE_MOVIE  = "1";
var CONST_MOVIE_SET     = "2";
var CONST_TV_EPISODES   = "3";

// Module Variables
var m_strLastTVShow     = "";
var m_strFileOption     = "";

// To include the common module
// File System Object
var m_objFSO = WScript.CreateObject("Scripting.FileSystemObject");

if (!m_objFSO) {
	WScript.Echo(L_InfoUnableToInclude_ErrorMessage + "\"Scripting.FileSystemObject\"");
	WScript.Quit(EXIT_METHOD_FAIL);
}

// Xml Document
var m_objXml = WScript.CreateObject("Microsoft.XMLDOM");

if (m_objXml == null) {
	WScript.Echo(L_InfoUnableToInclude_ErrorMessage + "\"Microsoft.XMLDOM\"");
	WScript.Quit(EXIT_METHOD_FAIL);
}

// Http request
var m_objReq = WScript.CreateObject("MSXML2.XMLHttp");

if (!m_objReq) {
	WScript.Echo(L_InfoUnableToInclude_ErrorMessage + "\"MSXML2.XMLHttp\"");
	WScript.Quit(EXIT_METHOD_FAIL);
}

// Run DOS command
var m_objApp = WScript.CreateObject("WScript.shell");

if (!m_objApp) {
	WScript.Echo(L_InfoUnableToInclude_ErrorMessage + "\"WScript.shell\"");
	WScript.Quit(EXIT_METHOD_FAIL);
}

// var strDetail = Download_Douban("25954475");
// WScript.Echo(strDetail);
// strDetail = strDetail.replace(/\n/g, "\\n");
// WScript.Echo(strDetail);
// var colDetail = eval("(" + strDetail + ")");
// for (var key in colDetail) {
//     JSON_Collection(colDetail[key], key);        
// }
//WScript.Quit();

// Calling the Main function
JSMain();

// end of the Main
WScript.Quit(EXIT_SUCCESS);


//********************************************************************
//* Function: JSMain
//*
//* Purpose: This is main function to starts execution
//*
//*
//* Input/ Output: None
//********************************************************************
function JSMain() {
	var strProcType = "";
	var strFileList = "";
	var strWorksheet = "";
	var strIOType = "";
	var strMessage = "";
	var intTimeCost = 0;
	var intRows = 0;
	var strArgs = "";
	var i, j, k, m, n, p, q, r;

// Get arguments
	if (WScript.Arguments.length < 1) {
		ShowHelpMessage();
	}

// Get argument 1 for version
	if (WScript.Arguments.Item(0).toUpperCase() == "-VERSION") {
		WScript.Echo(" Poster.js");
		WScript.Echo(" Version: " + CONST_Program_Version);
		WScript.Quit(EXIT_SUCCESS);
	}

	if (WScript.Arguments.length < 2) {
		ShowHelpMessage();
	}

// Get argument 1 to strProcType.
	var strProcType = WScript.Arguments.Item(0).toUpperCase();

// Get argument 2 to strProjPath
	var strProjPath = WScript.Arguments.Item(1);
    
    if (strProcType.toLowerCase().indexOf("_") > -1) {
        switch (strProcType) {
            case "DOUBAN_S":

                strProjPath = "https://api.douban.com/v2/movie/search?q=" + encodeURIComponent(strProjPath);
			
                break;
                
            case "DOUBAN_D":

                strProjPath = "https://api.douban.com/v2/movie/subject/" + strProjPath;
			
                break;
                
            case "MTIME_S":
            
                strProjPath = "http://m.mtime.cn/Service/callback.mi/Showtime/SearchVoice.api?keyword=" + encodeURIComponent(strProjPath);
                
                break; 
                
            case "MTIME_D":
            
                strProjPath = "http://m.mtime.cn/Service/callback.mi/movie/Detail.api?movieId=" + strProjPath;
                
                break;
                
            default:
            
                WScript.Echo(CONST_SEPARATE_LINE_1);
                WScript.Echo(CONST_LIST_LEVEL_1 + L_DeployTypeNotValid_ErrorMessage);
                WScript.Quit(EXIT_INVALID_PARAM);   
        }
        
        ResponseTest_JSON(strProjPath);
        
        return 0;    
    }
   
	if (strProjPath == ".") {
		strProjPath = m_objApp.CurrentDirectory;
	}

	if (strProjPath.lastIndexOf("\\") != strProjPath.length - 1) {
		strProjPath += "\\";
	}

	if (m_objFSO.FolderExists(strProjPath) == false) {
		WScript.Echo(CONST_LIST_LEVEL_1 + L_ProjectPathNotValid_ErrorMessage);
		WScript.Echo(CONST_LIST_LEVEL_1 + strProjPath);
		WScript.Echo(CONST_SEPARATE_LINE_1);
		WScript.Quit(EXIT_INVALID_PARAM);
	}

    // Get argument 3 to strFileType   
    if (WScript.Arguments.length > 2) {
        var strFileType = WScript.Arguments.Item(2);
    } else {
        var strFileType = CONST_SINGLE_MOVIE;
    }

	MOVIE_EXTENSION = "|" + MOVIE_EXTENSION + "|";

	// Start the process
	

	switch (strProcType) {
		case "TXT":
        
            DisplayResult(CONST_LIST_LEVEL_1 + "Searching text file in \"" + strProjPath + "\" ...", "", false, " ");

			intTimeCost = MovieSearch_Txt(strProjPath);

			break;

		case "MOV":
        case "DOUBAN":
        
            DisplayResult(CONST_LIST_LEVEL_1 + "Searching movie file in \"" + strProjPath + "\" ...", "", false, " ");
            
            intTimeCost = MovieSearch_Mov(strProjPath, strFileType, "douban");
			
			break;
            
        case "MTIME":
        
            DisplayResult(CONST_LIST_LEVEL_1 + "Searching movie file in \"" + strProjPath + "\" ...", "", false, " ");
            
            intTimeCost = MovieSearch_Mov(strProjPath, strFileType, "mtime");
			
			break;

		case "CLEAN":
            
            DisplayResult(CONST_LIST_LEVEL_1 + "Cleaning file in \"" + strProjPath + "\" ...", "", false, " ");
            
            WScript.Echo(CONST_SEPARATE_LINE_1);
            	
            intTimeCost = MovieSearch_Clean(strProjPath, 0, "");
            
			break;
            
        case "NEW":
        
            DisplayResult(CONST_LIST_LEVEL_1 + "Searching movie file in \"" + strProjPath + "\" ...", "", false, " ");
            
            intTimeCost = AddDateToFolder(strProjPath);
			
			break;

		default:

            WScript.Echo(CONST_SEPARATE_LINE_1);
			WScript.Echo(CONST_LIST_LEVEL_1 + L_DeployTypeNotValid_ErrorMessage);
			WScript.Quit(EXIT_INVALID_PARAM);

			break;
	}

// Output script runs seconds
	WScript.Echo(CONST_SEPARATE_LINE_1);

	i = intTimeCost / 1000;
	j = parseInt(i / 60.0);
	i = i - 60 * j;

	if (j > 0) {
		WScript.Echo(CONST_LIST_LEVEL_1 + "Process compeleted in " + j + ":" + i.toString().substring(0, i.toString().indexOf(".") + 4) +  " seconds.");
	} else {
		WScript.Echo(CONST_LIST_LEVEL_1 + "Process compeleted in " + i.toString().substring(0, i.toString().indexOf(".") + 4) +  " seconds.");
	}
}
//***************************  End of Main  **************************


//********************************************************************
//* Function: ShowHelpMessage
//*
//* Purpose: Display help message for this script.
//*
//*
//* Input: 	none.
//* Output: none.
//********************************************************************
function ShowHelpMessage() {
	WScript.Echo(EmptyLine_Text);
    WScript.Echo(L_ShowUsageLine01_Text);
    WScript.Echo(EmptyLine_Text);
    WScript.Echo(L_ShowUsageLine02_Text);
    WScript.Echo(L_ShowUsageLine03_Text);
    WScript.Echo(L_ShowUsageLine04_Text);
    WScript.Echo(L_ShowUsageLine05_Text);
    WScript.Quit(EXIT_INVALID_PARAM);
}


//********************************************************************
//* Function: MovieSearch_Clean
//*
//* Purpose: Clean *.nfo and *poster.jpg in specified folder.
//*
//*
//* Input:
//*		strFldMovie.
//* Output:
//* 	Time cost in second.
//*
//* Example:
//*   cscript Poster.js clean "D:\Movie"
//********************************************************************
function MovieSearch_Clean(strFldMovie, intFldLevel, strFldLevel) {
	var strNewLevel;
	var strFileName;
    var strEchoText;
	var i, j, k = 0, m = 0;
	var d1, d2;

	d1 = new Date();

	if (strFldMovie.lastIndexOf("\\") != strFldMovie.length - 1) {
		strFldMovie += "\\";
	}

	// 获取文件夹对象
	var objFolder = m_objFSO.GetFolder(strFldMovie);

	// 获取文件集合
	var enuFiles = new Enumerator(objFolder.Files);

	// 获取子文件夹集合
	var enuSubFolders = new Enumerator(objFolder.SubFolders);

	// 定义文件数组
	var aryMyFiles = new Array();

	// 获得文件夹中所有文件
	for (;!enuFiles.atEnd(); enuFiles.moveNext()) {
		var objFile = enuFiles.item();

		var strExtension = GetExtension(objFile.Name).toLowerCase();

		if (MOVIE_EXTENSION.indexOf("|" + strExtension + "|") > -1) {
			aryMyFiles.push(objFile.Name);
		}
	}

    if (aryMyFiles.length > 0) {
        // 对文件排序
    	aryMyFiles.sort();
    
    	// 依次处理文件
    	for (var i=0; i<aryMyFiles.length; i++) {
    		if (i < aryMyFiles.length - 1) {
    			strEchoText = CONST_LIST_LEVEL_1 + strFldLevel + "├";
                strNewLevel = CONST_LIST_LEVEL_1 + strFldLevel + "│";
    		} else {
    			strEchoText = CONST_LIST_LEVEL_1 + strFldLevel + "└";
                strNewLevel = CONST_LIST_LEVEL_1 + strFldLevel + "　"; 
    		}
            
            WScript.Echo(strEchoText + aryMyFiles[i]);
    
            // 去除文件扩展名
            strExtension = GetExtension(aryMyFiles[i]).toLowerCase();
            strFileName = aryMyFiles[i].replace("." + strExtension, "");
            
//             // 删除*.nfo文件
//         	if (m_objFSO.fileExists(strFldMovie + strFileName + ".nfo") == true) {
//                 m_objFSO.deleteFile(strFldMovie + strFileName + ".nfo");
//                 
//                 DisplayResult(strNewLevel + "└删除*.nfo文件", CONST_RESULT_OK, false);
//         	}
//             
//             // 删除tvshow.nfo文件
//         	if (m_objFSO.fileExists(strFldMovie + "tvshow.nfo") == true) {
//                 m_objFSO.deleteFile(strFldMovie + "tvshow.nfo");
//                 
//                 DisplayResult(strNewLevel + "└删除*.nfo文件", CONST_RESULT_OK, false);
//         	}
//             
            // 删除*-poster.jpg文件            
        	if (m_objFSO.fileExists(strFldMovie + strFileName + "-poster.jpg") == true) {
                m_objFSO.deleteFile(strFldMovie + strFileName + "-poster.jpg");
                
                DisplayResult(strNewLevel + "└删除*-poster.jpg文件", CONST_RESULT_OK, false);
        	}
//             
//             // 删除poster.jpg文件            
//         	if (m_objFSO.fileExists(strFldMovie + "poster.jpg") == true) {
//                 m_objFSO.deleteFile(strFldMovie + "poster.jpg");
//                 
//                 DisplayResult(strNewLevel + "└删除poster.jpg文件", CONST_RESULT_OK, false);
//         	}
//             // 将*-poster复制为poster.jpg文件
//             if (m_objFSO.fileExists(strFldMovie + strFileName + "-poster.jpg") == true) {
//                 m_objFSO.copyFile(strFldMovie + strFileName + "-poster.jpg"strFldMovie + "poster.jpg");
//                  
//                  DisplayResult(strNewLevel + "└复制为poster.jpg文件", CONST_RESULT_OK, false);    
//             }
//             // 将poster复制为*-poster.jpg文件            
//             if (m_objFSO.fileExists(strFldMovie + "poster.jpg") == true) {
//                 m_objFSO.copyFile(strFldMovie + "poster.jpg", strFldMovie + strFileName + "-poster.jpg");
//                  
//                  DisplayResult(strNewLevel + "└复制为*-poster.jpg文件", CONST_RESULT_OK, false);    
//             }
    	}
    }

	// 定义文件夹数组
	var aryMyFolders = new Array();

	// 获得文件夹中所有子文件夹
	for (;!enuSubFolders.atEnd(); enuSubFolders.moveNext()) {
		var objFolder = enuSubFolders.item();

		aryMyFolders.push(objFolder.Name);
	}

	// 对子文件夹排序
	aryMyFolders.sort();

	// 若存在子文件夹
	if (aryMyFolders.length > 0) {
		for (var i=0; i<aryMyFolders.length; i++) {
			if (i < aryMyFolders.length - 1) {
				WScript.Echo(CONST_LIST_LEVEL_1 + strFldLevel + "├" + aryMyFolders[i]);

				strNewLevel = strFldLevel + "│";
			} else {
				WScript.Echo(CONST_LIST_LEVEL_1 + strFldLevel + "└" + aryMyFolders[i]);

				strNewLevel = strFldLevel + "　";
			}

			if (aryMyFolders[i].substring(0,1) != ".") {
				// 递归查找子文件
				MovieSearch_Clean(strFldMovie + aryMyFolders[i], intFldLevel + 1, strNewLevel);
			}
		}
	} else {
		// 且未找到视频文件
		if (aryMyFiles.length < 1) {
			DisplayResult(CONST_LIST_LEVEL_1 + strFldLevel + "└File not found ", CONST_RESULT_FAILED, false);
		}
	}

	d2 = new Date();

	return(d2.valueOf() - d1.valueOf());
}


//********************************************************************
//* Function: MovieSearch_Txt
//*
//* Purpose: Search *.txt file in specified folder.
//*
//*
//* Input:
//*		strFldMovie.
//* Output:
//* 	Time cost in second.
//*
//* Example:
//*   cscript Poster.js txt "D:\Movie"
//********************************************************************
function MovieSearch_Txt(strFldMovie) {
	var strNewLevel;
	var strFileName;
	var strExtension;
	var strMovieFile;
	var i, j, k = 0, m = 0;
	var d1, d2;

	d1 = new Date();

	if (strFldMovie.lastIndexOf("\\") != strFldMovie.length - 1) {
		strFldMovie += "\\";
	}

	// 获取文件夹对象
	var objFolder = m_objFSO.GetFolder(strFldMovie);

	// 获取文件集合
	var enuFiles = new Enumerator(objFolder.Files);

	// 定义文件数组
	var aryMyFiles = new Array();

	// 获得文件夹中所有文件
	for (;!enuFiles.atEnd(); enuFiles.moveNext()) {
		var objFile = enuFiles.item();
        
        strFileName = objFile.Name;
        
        if (strFileName.toLowerCase() == "thumbs.db") {
			try {
				m_objFSO.deleteFile(strFilePath + strFndFile, true);
			} catch (e) {}	
		}

		strExtension = GetExtension(objFile.Name).toLowerCase();

		if (strExtension == "txt") {
			// 记录该文件
            aryMyFiles.push(objFile.Name);
		}
	}

	// 对文件排序
	aryMyFiles.sort();

	// 依次处理文件
	for (var i=0; i<aryMyFiles.length; i++) {
        WScript.Echo(CONST_SEPARATE_LINE_1.replace(/-/g, "="));
        
		DisplayResult(CONST_LIST_LEVEL_1 + "发现目标文件：" + aryMyFiles[i], "", false, " ");

		// 处理该文件
		CreateNfo_TxtFile(strFldMovie, aryMyFiles[i]);
	}

	// 未找到目标文件
	if (aryMyFiles.length < 1) {
		DisplayResult(CONST_LIST_LEVEL_1 + "发现目标文件：", CONST_RESULT_FAILED, false);
	}

	d2 = new Date();

	return(d2.valueOf() - d1.valueOf());
}


function CreateNfo_TxtFile(strFilePath, strTxtFile) {
    var strField;
    var strValue, strTitle;
    var intCharAt;
    var strPart = "";
    var strPlot = "";
	var strName, strRole;
	var aryValue;

	if (strFilePath.lastIndexOf("\\") != strFilePath.length - 1) {
		strFilePath += "\\";
	}

	if (m_objFSO.FileExists(strFilePath + strTxtFile) == false) {
		return false;
	}

	m_objXml.async = true;

    var objTxt = m_objFSO.OpenTextFile(strFilePath + strTxtFile, CONST_ForReading , false, 0);

    m_objXml.loadXML("<?xml version=\"1.0\" encoding=\"utf-8\" ?><movie />");

	var eleMovie = m_objXml.documentElement;

	if (strFilePath.indexOf("系列：") > -1) {
		aryValue = strFilePath.split("\\");

		// 获取倒数第二层文件夹名称
		strValue = aryValue[aryValue.length - 3];

		// 添加集合名称
		Xml_AddNode(m_objXml, eleMovie, "set", strValue.replace("[", "").replace("]", ""));

		// 获取倒数第一层文件夹名称
		strValue = aryValue[aryValue.length - 2];
		aryValue = strValue.split("][");

        // 获取中文名称
		strTitle = aryValue[1]

		// 添加分集名称
		Xml_AddNode(m_objXml, eleMovie, "sorttitle", strTitle);
	}

    while (!objTxt.AtEndOfStream) {
        strRead = objTxt.ReadLine().replace(/(^\s*)|(\s*$)/g, "");

		if (strRead == "") {
			continue;
		}

		strField = strRead.substring(0, 5);
        strValue = strRead.substring(6, strRead.length).replace(/(^\s*)|(\s*$)/g, "");

        switch (strField) {
        	case "◎译　　名":
            case "◎中 文 名":

        		strPart = "";

        		Xml_AddNode(m_objXml, eleMovie, "title", strValue);

        		break;

			case "◎片　　名":
            case "◎英 文 名":

        		strPart = "";

        		Xml_AddNode(m_objXml, eleMovie, "originaltitle", strValue);

        		break;	　

			case "◎年　　代":

        		strPart = "";

        		Xml_AddNode(m_objXml, eleMovie, "year", strValue);

        		break;	　

        	case "◎国　　家":

        		strPart = "";

                aryValue = strValue.split("/");
                
                for (var i = 0; i < aryValue.length; i++) {
                    Xml_AddNode(m_objXml, eleMovie, "country", aryValue[i]);
                }
        		

        		break;	　

        	case "◎类　　别":

        		strPart = "";

        		aryValue = strValue.split("/");
                
                for (var i = 0; i < aryValue.length; i++) {
                    Xml_AddNode(m_objXml, eleMovie, "genre", aryValue[i]);
                }

        		break;	　

        	case "◎IMDB":

        		strPart = "";

        		strField = strRead.substring(0, 7);
        		strValue = strRead.substring(8, strRead.length).replace(/(^\s*)|(\s*$)/g, "");

        		if (strField == "◎IMDB评分") {
                    aryValue = strValue.split("/");
                    
					Xml_AddNode(m_objXml, eleMovie, "rating", aryValue[0]);
				}

        		if (strField == "◎IMDB链接") {
        			intCharAt = strValue.indexOf("/tt");

	        		if (intCharAt > -1) {
	        			strValue = strValue.substring(intCharAt + 1, strValue.length);
                      
                        intCharAt = strValue.indexOf("[");
                        
                        if (intCharAt < 0) {
                            intCharAt = strValue.length   
                        }
                        
                        strValue = strValue.substring(0, intCharAt);
                        
	        			Xml_AddNode(m_objXml, eleMovie, "id", strValue);
	        		}
        		}

        		break;	　

        	case "◎导　　演":

        		strPart = "director";

        		Xml_AddNode(m_objXml, eleMovie, "director", strValue);

        		break;	　

        	case "◎主　　演":

        		strPart = "actor";

        		aryValue = strValue.split(".");
                
                var actor = Xml_AddNode(m_objXml, eleMovie, "actor", "");
                
        		strName = aryValue[0].replace(/(^\s*)|(\s*$)/g, "");
        		                
				Xml_AddNode(m_objXml, actor, "name", strName);
                
                if (aryValue.length > 5) {
                    strRole = aryValue[aryValue.length - 1].replace(/(^\s*)|(\s*$)/g, "");
                    
                    Xml_AddNode(m_objXml, actor, "role", strRole);
                }

        		break;	　

        	case "◎简　　介":

        		strPart = "plot";

        		strPlot = strPlot + strValue;

        		break;	　

        	default:

        		switch (strPart) {
					case "director":

						Xml_AddNode(m_objXml, eleMovie, "director", strRead);

						break;

					case "actor":
                    
						var actor = Xml_AddNode(m_objXml, eleMovie, "actor", "");
                        
                        aryValue = strRead.split(".");

		        		strName = aryValue[0].replace(/(^\s*)|(\s*$)/g, "").replace(/　/g, "");     						
                        
						Xml_AddNode(m_objXml, actor, "name", strName);
                        
						if (aryValue.length > 5) {
                            strRole = aryValue[aryValue.length - 1].replace(/(^\s*)|(\s*$)/g, "");
                            
                            Xml_AddNode(m_objXml, actor, "role", strRole);
                        }

						break;

					case "plot":

                        if (strRead.indexOf("幕后制作") > -1 || strRead.indexOf("一句话评论") > -1) {
                            strPart = "";
                        } else {
                            strPlot = strPlot + "\n" + strRead;
                        }
                        
						break;

					default:

		        		break;
				}

        		break;
        }
    }

	Xml_AddNode(m_objXml, eleMovie, "plot", strPlot);

    objTxt.Close();

    objFSO = null;

    m_objXml.save(strFilePath + strTxtFile.replace(".txt", ".nfo"));
    
    DisplayResult(CONST_LIST_LEVEL_1 + "生成文件：*.nfo", CONST_RESULT_OK, false);

    return true;
}


function ResponseTest_JSON(strQueryUrl) {
    WScript.Echo(strQueryUrl);
    WScript.Echo(CONST_SEPARATE_LINE_1);
     
    m_objReq.open("GET", strQueryUrl, false);
	m_objReq.send();

    try {
        var colMsg = eval("(" + m_objReq.responseText + ")");
    } catch (e) {
        WScript.Echo(CONST_LIST_LEVEL_1 + "Response Error: " + strQueryUrl);
        
        return false; 
    }
       
    if (colMsg.code) {
    	WScript.Echo(CONST_LIST_LEVEL_1 + "code: " + colMsg.count);
    	WScript.Echo(CONST_LIST_LEVEL_1 + "msg:  " + colMsg.msg);
    	WScript.Echo(CONST_LIST_LEVEL_1 + "request: " + colMsg.request);
    
    	return false;
    }
    
    for (var key in colMsg) {
        JSON_Collection(colMsg[key], key);        
    }
        
    return true;   
}


function JSON_Collection(objJSON, strEcho) {    
    if (objJSON instanceof Array) {
        for (var i = 0; i < objJSON.length; i++) {
            JSON_Collection(objJSON[i], strEcho + "[" + i + "]");
        }
        
        return 0;  
    }
    
    if (objJSON instanceof Object) {
        for (var key in objJSON) {
            JSON_Collection(objJSON[key], strEcho + "." + key);
        } 
        
        return 0; 
    } 
    
    if (objJSON == null) {
        WScript.Echo(strEcho + " = null");
    } else {
        WScript.Echo(strEcho + " = \"" + objJSON + "\"");
    }     
}


//********************************************************************
//* Function: MovieSearch_Mov
//*
//* Purpose: Search movie in specified path and download from douban.
//*
//*
//* Input:
//*		strFldMovie    Path of movie folder
//*     strSearchIn    douban
//* Output:
//* 	Time cost in second.
//*
//* Example:
//*   cscript Poster.js new .
//********************************************************************
function AddDateToFolder(strFilePath) {
    var objFolder = m_objFSO.GetFolder(strFilePath);
    var colFolder = new Enumerator(objFolder.SubFolders);
    
    var d1 = new Date();

    
    for (; !colFolder.atEnd(); colFolder.moveNext()) {
    	var strName = colFolder.item().Name;
    	var strDate = colFolder.item().DateCreated;
        
    	strDate = FormatDate(strDate);
    
    	if (strName.indexOf(strDate) > -1) {
    		continue;
    	}
    
        WScript.Echo(CONST_SEPARATE_LINE_1.replace(/-/g, "="));
        
    	try {
    		m_objFSO.moveFolder(strFilePath + strName, strFilePath + strDate + "." + strName);
    
    		DisplayResult(CONST_LIST_LEVEL_1 + "重命名文件夹：" + strDate + "." + strName + " ", CONST_RESULT_OK, false);
            
            while (m_objFSO.folderExists(strFilePath + strDate + "." + strName) == false) {
                WScript.sleep(PAUSE_DURATION);    
            }
            
            MovieSearch_Mov(strFilePath + strDate + "." + strName, "1", "douban");
            
    	} catch (e) {            
    		DisplayResult(CONST_LIST_LEVEL_1 + "重命名文件夹：" + strDate + "." + strName + " ", CONST_RESULT_FAILED, false);
            WScript.Echo(CONST_LIST_LEVEL_1 + e.toString());
    	}
    }
    
    var d2 = new Date();

	return(d2.valueOf() - d1.valueOf());
}

//********************************************************************
//* Function: MovieSearch_Mov
//*
//* Purpose: Search movie in specified path and download from douban.
//*
//*
//* Input:
//*		strFldMovie    Path of movie folder
//*     strSearchIn    douban
//* Output:
//* 	Time cost in second.
//*
//* Example:
//*   cscript Poster.js Mov "D:\Movie"
//********************************************************************
function MovieSearch_Mov(strFldMovie, strFileType, strSearchIn) {
	var strFileName;
	var strExtension;
	var strMovieFile;
	var i, j, k = 0, m = 0;
	var d1, d2;

	d1 = new Date();

	if (strFldMovie.lastIndexOf("\\") != strFldMovie.length - 1) {
		strFldMovie += "\\";
	}

	// 获取文件夹对象
	var objFolder = m_objFSO.GetFolder(strFldMovie);

	// 获取文件集合
	var enuFiles = new Enumerator(objFolder.Files);

	// 获取子文件夹集合
	var enuSubFolders = new Enumerator(objFolder.SubFolders);

	// 定义文件数组
	var aryMyFiles = new Array();

	// 获得文件夹中所有文件
	for (;!enuFiles.atEnd(); enuFiles.moveNext()) {
		var objFile = enuFiles.item();

		strExtension = GetExtension(objFile.Name).toLowerCase();

		if (MOVIE_EXTENSION.indexOf("|" + strExtension + "|") > -1) {
			// 记录该文件
            aryMyFiles.push(objFile.Name);
		}
	}

	// 对文件排序
	aryMyFiles.sort();

	// 依次处理文件
	for (var i=0; i<aryMyFiles.length; i++) {
		// 处理该文件
        Download_Detail(strSearchIn, strFldMovie, aryMyFiles[i], strFileType);		
	}

	// 定义文件夹数组
	var aryMyFolders = new Array();

	// 获得文件夹中所有子文件夹
	for (;!enuSubFolders.atEnd(); enuSubFolders.moveNext()) {
		var objFolder = enuSubFolders.item();

		aryMyFolders.push(objFolder.Name);
	}

	// 对子文件夹排序
	aryMyFolders.sort();

	// 若存在子文件夹
	if (aryMyFolders.length > 0) {
		for (var i=0; i<aryMyFolders.length; i++) {
			if (aryMyFolders[i].substring(0,1) != ".") {
				// 递归查找子文件
				MovieSearch_Mov(strFldMovie + aryMyFolders[i], strFileType, strSearchIn);
			}
		}
	} 
    
	d2 = new Date();

	return(d2.valueOf() - d1.valueOf());
}


function Download_Detail(strSearchIn, strFilePath, strFileName, strFileType) {
	var strQueryTxt = GetMovieKeyword(strFileName).replace(/e[p]?[0-9]{1,2}/gi, "");
    var strExtension = GetExtension(strFileName);
    var strMovTitle = "";
    var strGoOption = "";
    
    var strFile_Nfo = strFileName.replace("." + strExtension, ".nfo");
    var strFile_Pst = "poster.jpg";
    
    if (strFileType == "3") {
        if (m_strLastTVShow != strFilePath) {
            m_strLastTVShow = strFilePath; 
            m_strFileOption = "a";
            
            WScript.Echo(CONST_SEPARATE_LINE_1.replace(/-/g, "="));
        
            DisplayResult(CONST_LIST_LEVEL_1 + "发现剧集影片：" + strQueryTxt, "", false, " ");           
        } else {
            if (m_strFileOption == "0" || (m_strFileOption != "a" && m_strFileOption != "s")) {
        		return false;
        	}
                    
            // 生成剧集的episodename.nfo文件
            return CreateNof_TVEpisodes(strFilePath, strFileName, m_strFileOption);
        }
        
        var strFile_Nfo = "tvshow.nfo";
        var strFile_Pst = "poster.jpg";
    } else {
        WScript.Echo(CONST_SEPARATE_LINE_1.replace(/-/g, "="));
        
        DisplayResult(CONST_LIST_LEVEL_1 + "发现影片文件：" + strFileName, "", false, " ");
    }
    
    if (strFilePath.lastIndexOf("\\") != strFilePath.length - 1) {
		strFilePath += "\\";
	}
       
    var blnNfoExist = m_objFSO.fileExists(strFilePath + strFile_Nfo);       
    var blnPstExist = m_objFSO.fileExists(strFilePath + strFile_Pst);
    var blnBothFile = false;
    
    if (blnNfoExist == true || blnPstExist == true) {
        WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);
        
        strGoOption = AskForOption(CONST_LIST_LEVEL_1 + "文件已存在，是否替换？(a=全部替换; s=单独替换)：");
        
        m_strFileOption = strGoOption;
        
        if (strGoOption.toLowerCase() != "a" && strGoOption.toLowerCase() != "s") {
            if (strFileType == "3") {
                DisplayResult(CONST_LIST_LEVEL_1 + "剧集影片：" + strQueryTxt, CONST_RESULT_IGNORE, false);
            } else {
                DisplayResult(CONST_LIST_LEVEL_1 + "生成文件：*.nfo", CONST_RESULT_IGNORE, false);
                DisplayResult(CONST_LIST_LEVEL_1 + "下载海报：poster.jpg", CONST_RESULT_IGNORE, false);
            }
            
            return false;
        } 
        
        if (strGoOption.toLowerCase() == "a") {
            blnBothFile = true;
            blnNfoExist = false;
            blnPstExist = false;
        }
    }
    
	var strMovieKey = DisplaySearching(strSearchIn, strQueryTxt);

	if (strMovieKey == "") {
        if (strFileType == "3") {
            DisplayResult(CONST_LIST_LEVEL_1 + "剧集影片：" + strQueryTxt, CONST_RESULT_IGNORE, false);
        } else {
            DisplayResult(CONST_LIST_LEVEL_1 + "生成文件：*.nfo", CONST_RESULT_IGNORE, false);
            DisplayResult(CONST_LIST_LEVEL_1 + "下载海报：poster.jpg", CONST_RESULT_IGNORE, false);
        }

		return false;
	}
    
    switch (strSearchIn.toLowerCase()) {
        case "douban" :
       
            var strDetail = Download_Douban(strMovieKey);
        
            break;
            
        case "mtime" :
        
            var strDetail = Download_Mtime(strMovieKey);
        
            break;
            
        default :
        
            return false;
    }
    
    var colDetail = eval("(" + strDetail + ")");
    
    if (colDetail.err) {
        DisplayResult(CONST_LIST_LEVEL_1 + "获取影片信息：" + strMovieKey, CONST_RESULT_FAILED, false);
        
        WScript.Echo(CONST_LIST_LEVEL_1 + "err: " + colDetail.err);
		WScript.Echo(CONST_LIST_LEVEL_1 + "msg: " + colDetail.msg);
        
        return false;
    }
    
    if (blnBothFile == false && blnNfoExist == true) {
        WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);
        
        strGoOption = AskForOption(CONST_LIST_LEVEL_1 + "NFO文件已存在，是否替换？(y=替换)：");
        
        if (strGoOption.toLowerCase() == "y") {
            blnNfoExist = false;
        } else {    
            DisplayResult(CONST_LIST_LEVEL_1 + "生成文件：*.nfo", CONST_RESULT_IGNORE, false);
        } 
    }
        
    if (blnNfoExist == false) {    
		m_objXml.async = true;

        switch (strFileType) {
            case "1":
           
                m_objXml.loadXML("<?xml version=\"1.0\" encoding=\"utf-8\" ?><movie />");
    
        		var eleMovie = m_objXml.documentElement;
                
                if (strFilePath.indexOf("系列：") > -1) {
                    Xml_AddNode(m_objXml, eleMovie, "set", GetMovieSet_Set(strFilePath + strFileName));
                    Xml_AddNode(m_objXml, eleMovie, "sorttitle", GetMovieSet_Sorttitle(strFilePath + strFileName));
                }
                
        		Xml_AddNode(m_objXml, eleMovie, "title", GetMovie_Title(strFilePath + strFileName, colDetail.title));
                Xml_AddNode(m_objXml, eleMovie, "originaltitle", colDetail.originaltitle);
                Xml_AddNode(m_objXml, eleMovie, "year", colDetail.year);            		
                Xml_AddNode(m_objXml, eleMovie, "rating", colDetail.rating);
                
                for (var i = 0; i < colDetail.countries.length; i++) {
                    Xml_AddNode(m_objXml, eleMovie, "country", colDetail.countries[i]);
                }
                
                for (var i = 0; i < colDetail.genres.length; i++) {
                    Xml_AddNode(m_objXml, eleMovie, "genre", colDetail.genres[i]);
        		}
        
        		for (var i = 0; i < colDetail.directors.length; i++) {
                    Xml_AddNode(m_objXml, eleMovie, "director", colDetail.directors[i].name);	
        		}
        
        		for (var i = 0; i < colDetail.actors.length; i++) {
        			var objNewNode = Xml_AddNode(m_objXml, eleMovie, "actors", "");
        
        			Xml_AddNode(m_objXml, objNewNode, "name", colDetail.actors[i].name);
                    
                    if (colDetail.actors[i].role) {
                        Xml_AddNode(m_objXml, objNewNode, "role", colDetail.actors[i].role);
                    }
                    
                    if (colDetail.actors[i].thumb) {
                        Xml_AddNode(m_objXml, objNewNode, "thumb", colDetail.actors[i].thumb);
                    }
        		}
        
                Xml_AddNode(m_objXml, eleMovie, "id", colDetail.id);
        		Xml_AddNode(m_objXml, eleMovie, "plot", colDetail.plot);
                
                m_objXml.save(strFilePath + strFile_Nfo);

		        DisplayResult(CONST_LIST_LEVEL_1 + "生成文件：*.nfo", CONST_RESULT_OK, false);
                
                break;
            
            case "2":
            
                m_objXml.loadXML("<?xml version=\"1.0\" encoding=\"utf-8\" ?><movie />");
    
        		var eleMovie = m_objXml.documentElement;
                
                Xml_AddNode(m_objXml, eleMovie, "set", GetMovieSet_Set(strFilePath + strFileName));
                Xml_AddNode(m_objXml, eleMovie, "sorttitle", GetMovieSet_Sorttitle(strFilePath + strFileName));
                
        		Xml_AddNode(m_objXml, eleMovie, "title", GetMovie_Title(strFilePath + strFileName, colDetail.title)); 
                Xml_AddNode(m_objXml, eleMovie, "originaltitle", colDetail.original_title);
                Xml_AddNode(m_objXml, eleMovie, "year", colDetail.year);                    
        		Xml_AddNode(m_objXml, eleMovie, "rating", colDetail.rating);
                
                for (var i = 0; i < colDetail.countries.length; i++) {
                    Xml_AddNode(m_objXml, eleMovie, "country", colDetail.countries[i]);
                }
                
                for (var i = 0; i < colDetail.genres.length; i++) {
                    Xml_AddNode(m_objXml, eleMovie, "genre", colDetail.genres[i]);
        		}
        
        		for (var i = 0; i < colDetail.directors.length; i++) {
                    Xml_AddNode(m_objXml, eleMovie, "director", colDetail.directors[i].name);	
        		}
        
        		for (var i = 0; i < colDetail.actors.length; i++) {
        			var objNewNode = Xml_AddNode(m_objXml, eleMovie, "actor", "");
        
        			Xml_AddNode(m_objXml, objNewNode, "name", colDetail.actors[i].name);
                    
                    if (colDetail.actors[i].thumb) {
                        Xml_AddNode(m_objXml, objNewNode, "thumb", colDetail.actors[i].thumb);
                    }
        		}
        
                Xml_AddNode(m_objXml, eleMovie, "id", colDetail.id);
        		Xml_AddNode(m_objXml, eleMovie, "plot", colDetail.plot);
                
                m_objXml.save(strFilePath + strFile_Nfo);

		        DisplayResult(CONST_LIST_LEVEL_1 + "生成文件：*.nfo", CONST_RESULT_OK, false);
                
                break;
                
            case "3":
                // 生成剧集的tvshow.nfo文件
                m_objXml.loadXML("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\" ?><tvshow />");
    
        		var eleMovie = m_objXml.documentElement;
                
                var strTitle = GetMovie_Title(strFilePath + strFileName, colDetail.title);
                var strSeason = GetTVEpisodes_Season(strFilePath + strFileName);

                if (strSeason != "0") {
                    strTitle = strTitle + ".第" + strSeason + "季";    
                } 
                
                Xml_AddNode(m_objXml, eleMovie, "set", colDetail.original_title);            
        		Xml_AddNode(m_objXml, eleMovie, "title", strTitle);
                Xml_AddNode(m_objXml, eleMovie, "showtitle", strTitle);
                Xml_AddNode(m_objXml, eleMovie, "rating", colDetail.rating);
                
                Xml_AddNode(m_objXml, eleMovie, "year", colDetail.year);
                Xml_AddNode(m_objXml, eleMovie, "aired", colDetail.year);                    
                
                Xml_AddNode(m_objXml, eleMovie, "season", colDetail.season);
        		Xml_AddNode(m_objXml, eleMovie, "episode", colDetail.episode);                 
        		Xml_AddNode(m_objXml, eleMovie, "studio", GetTVEpisodes_Studio(strFilePath + strFileName));
                
                for (var i = 0; i < colDetail.countries.length; i++) {
                    Xml_AddNode(m_objXml, eleMovie, "country", colDetail.countries[i]);
                }
                
                for (var i = 0; i < colDetail.genres.length; i++) {
                    Xml_AddNode(m_objXml, eleMovie, "genre", colDetail.genres[i]);
        		}

        		for (var i = 0; i < colDetail.directors.length; i++) {
        			Xml_AddNode(m_objXml, eleMovie, "director", colDetail.directors[i].name);
        		}
        
        		for (var i = 0; i < colDetail.actors.length; i++) {
        			var objNewNode = Xml_AddNode(m_objXml, eleMovie, "actor", "");
        
        			Xml_AddNode(m_objXml, objNewNode, "name", colDetail.actors[i].name);
                    
                    if (colDetail.actors[i].thumb) {
                        Xml_AddNode(m_objXml, objNewNode, "thumb", colDetail.actors[i].thumb);
                    }
        		}
        
        		Xml_AddNode(m_objXml, eleMovie, "plot", colDetail.plot);
                
                m_objXml.save(strFilePath + strFile_Nfo);
                
                DisplayResult(CONST_LIST_LEVEL_1 + "生成文件：tvshow.nfo", CONST_RESULT_OK, false);
                
                break;
                
            default:
            
                DisplayResult(CONST_LIST_LEVEL_1 + "生成文件：" + strFile_Nfo, CONST_RESULT_FAILED, false);
                
                break;
        }
    }
        
    if (blnBothFile == false && blnPstExist == true) {        
        var strGoOption = AskForOption(CONST_LIST_LEVEL_1 + "JPG文件已存在，是否替换？(y=替换)：");
        
        if (strGoOption.toLowerCase() == "y") {
            blnPstExist = false;
        } else {    
            DisplayResult(CONST_LIST_LEVEL_1 + "下载海报：poster.jpg", CONST_RESULT_IGNORE, false);
        } 
    }
        
    if (blnPstExist == false) {
		if (colDetail.poster) {
            m_objApp.run('"' + m_objFSO.GetFile(WScript.ScriptFullName).ParentFolder.Path + '\\wget.exe" "' + colDetail.poster + '" -O "' + strFilePath + strFile_Pst + '"', true);
            
            DisplayResult(CONST_LIST_LEVEL_1 + "下载海报：poster.jpg", CONST_RESULT_OK, false);            
        } else {
            DisplayResult(CONST_LIST_LEVEL_1 + "下载海报：poster.jpg", CONST_RESULT_FAILED, false);    			
		} 
    }
    
    if (strFileType == "3") {            
        // 生成剧集的episodename.nfo文件
        CreateNof_TVEpisodes(strFilePath, strFileName, m_strFileOption);
    }
    
    WScript.sleep(PAUSE_DURATION);

	return true;
}


function Download_Douban(strMovieKey) {
    var strJsonText = "";
    var strQueryUrl = "https://api.douban.com/v2/movie/subject/" + strMovieKey;

	m_objReq.open("GET", strQueryUrl, false);
	m_objReq.send();

	try {
		var colMsg = eval("(" + m_objReq.responseText + ")");

		if (colMsg.code) {
            strJsonText += '"err": "' + colMsg.code + '", ';
            strJsonText += '"msg": "' + colMsg.msg + '"';
            
			return '{' + strJsonText + '}';
		}

        strJsonText += '"title": "' + colMsg.title + '", ';
        strJsonText += '"originaltitle": "' + colMsg.original_title + '", ';
        strJsonText += '"year": "' + colMsg.year + '", ';
        strJsonText += '"rating": "' + colMsg.rating.average + '", ';
        
        strJsonText += '"countries": ["' + colMsg.countries.join('", "') + '"], ';
        strJsonText += '"genres": ["' + colMsg.genres.join('", "') + '"], ';
        
        strJsonText += '"directors": [';
           
		for (var i = 0; i < colMsg.directors.length; i++) {
            var strDirector = colMsg.directors[i].name;
            var aryDirector = strDirector.split(",");
            
            if (i > 0) {
                strJsonText += ', ';
            }
            
            for (var j = 0; j < aryDirector.length; j++) {
                strJsonText += '{"name": "' + aryDirector[j].replace(/(^\s*)|(\s*$)/g, "") + '", "thumb": null}';
            }	
		}
        
        strJsonText += '], ';
        
        strJsonText += '"actors": [';
            
		for (var i = 0; i < colMsg.casts.length; i++) {
			if (i > 0) {
                strJsonText += ', ';
            }

            strJsonText += '{"name": "' + colMsg.casts[i].name + '", ';
            strJsonText += '"role": null, ';
            
            if (colMsg.casts[i].avatars) {
                strJsonText += '"thumb": "' + colMsg.casts[i].avatars.large.replace(/\//g, "\/").replace("https:", "http:") + '"}';
            } else {
                strJsonText += '"thumb": null}';
            }
		}
        
        strJsonText += '], ';
        
        strJsonText += '"id": "' + colMsg.id + '", ';
        strJsonText += '"plot": "' + colMsg.summary.replace(/\n/g, "\\n").replace(/"/g, "\\\"") + '", ';
        
        strJsonText += '"poster": "' + colMsg.images.large.replace(/\//g, "\/").replace("https:", "http:") + '", ';
        strJsonText += '"fanart": null, ';
        
        if (colMsg.current_season == null) {
            strJsonText += '"season": "0", ';
        } else {
            strJsonText += '"season": "' + colMsg.current_season + '", ';
        }
        
        if (colMsg.episodes_count == null) {
            strJsonText += '"episode": "0" ';
        } else {
            strJsonText += '"episode": "' + colMsg.episodes_count + '" ';
        }
        
        return '{' + strJsonText + '}';

	} catch (e) {
		strJsonText += '"err": "0000", ';
        strJsonText += '"msg": "' + e.toString() + '"';
            
		return '{' + strJsonText + '}';
	}

	return '{' + strJsonText + '}';
}


function Download_Mtime(strMovieKey) {
    var strJsonText = "";
    var strQueryUrl = "http://m.mtime.cn/Service/callback.mi/movie/Detail.api?movieId=" + strMovieKey;

	m_objReq.open("GET", strQueryUrl, false);
	m_objReq.send();

	try {
		var colMsg = eval("(" + m_objReq.responseText + ")");

		if (colMsg.code) {
            strJsonText += '"err": "' + colMsg.code + '", ';
            strJsonText += '"msg": "' + colMsg.msg + '"';
            
			return '{' + strJsonText + '}';
		}

        strJsonText += '"title": "' + colMsg.titleCn + '", ';
        strJsonText += '"originaltitle": "' + colMsg.titleEn + '", ';
        strJsonText += '"year": "' + colMsg.year + '", ';
        strJsonText += '"rating": "' + colMsg.rating + '", ';
        
        strJsonText += '"countries": ["' + colMsg.release.location + '"], ';
        strJsonText += '"genres": ["' + colMsg.type.join('", "') + '"], ';
        
        strJsonText += '"directors": [';
        
        if (colMsg.director instanceof Array) {   
    		for (var i = 0; i < colMsg.director.length; i++) {
                if (i > 0) {
                    strJsonText += ', ';
                }
                
                strJsonText += '{"name": "' + colMsg.director[i].directorName + '\/' + colMsg.director[i].directorNameEn + '", ';
                
                if (colMsg.director[i].directorImg) {
                    strJsonText += '"thumb": "' + colMsg.director[i].directorImg.replace(/\//g, "\/") + '"}';
                } else {
                    strJsonText += '"thumb": null}';
                }	
    		}
        } else {
            strJsonText += '{"name": "' + colMsg.director.directorName + '\/' + colMsg.director.directorNameEn + '", ';
                
            if (colMsg.director.directorImg) {
                strJsonText += '"thumb": "' + colMsg.director.directorImg.replace(/\//g, "\/") + '"}';
            } else {
                strJsonText += '"thumb": null}';
            }
        }
        
        strJsonText += '], ';
        
        strJsonText += '"actors": [';
            
		for (var i = 0; i < colMsg.actorList.length; i++) {
			if (i > 0) {
                strJsonText += ', ';
            }

            strJsonText += '{"name": "' + colMsg.actorList[i].actor + '\/' + colMsg.actorList[i].actorEn + '", ';
            strJsonText += '"role": "' + colMsg.actorList[i].roleName + '", ';
            
            if (colMsg.actorList[i].actorImg) {
                strJsonText += '"thumb": "' + colMsg.actorList[i].actorImg.replace(/\//g, "\/") + '"}';
            } else {
                strJsonText += '"thumb": null}';
            }
		}
        
        strJsonText += '], ';
        
        strJsonText += '"id": "' + strMovieKey + '", ';
        strJsonText += '"plot": "' + colMsg.content.replace(/\n/g, "\\n") + '", ';
        
        strJsonText += '"poster": "' + colMsg.image.replace(/\//g, "\/") + '", ';
        strJsonText += '"fanart": null, ';
        
        strJsonText += '"season": "0", ';
        strJsonText += '"episode": "0" ';
                
        return '{' + strJsonText + '}';

	} catch (e) {
		strJsonText += '"err": "0000", ';
        strJsonText += '"msg": "' + e.toString() + '"';
            
		return '{' + strJsonText + '}';
	}

	return '{' + strJsonText + '}';
}


// 以列表方式显示搜索结果，并提供选择
function DisplaySearching(strSearchIn, strQueryTxt) {
    var arySearchRt = new Array();
    
	WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);
    
	DisplayResult(CONST_LIST_LEVEL_1 + "正在搜索影片：\"" + strQueryTxt + "\" from " + strSearchIn.toLowerCase(), "", false, " ");

    switch (strSearchIn.toLowerCase()) {
        case "douban":
        
            arySearchRt = Searching_Douban(strQueryTxt);
                                                      
            break;
            
        case "mtime":
            
            arySearchRt = Searching_Mtime(strQueryTxt);
            
            break;
            
    }

    WScript.Echo(CONST_LIST_LEVEL_1 + "找到影片数目：" + arySearchRt.length);
	WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);

    var blnGetAgain = false;
        
	if (arySearchRt.length < 1) {
		blnGetAgain = true;
	}
        
    while (blnGetAgain == true) {
        var strOptionRt = AskForOption(CONST_LIST_LEVEL_1 + "是否重新输入影片关键字进行查找？(y=重新搜索)：");

		if (strOptionRt != "y") {
			WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);
            
            m_strFileOption = "0";

			return "";
		}
        
        strQueryTxt = AskForOption(CONST_LIST_LEVEL_1 + "请重新输入影片关键字：");
        
        WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);

    	DisplayResult(CONST_LIST_LEVEL_1 + "正在搜索影片：\"" + strQueryTxt + "\" from " + strSearchIn.toLowerCase(), "", false, " ");
    
    	switch (strSearchIn.toLowerCase()) {
            case "douban":
            
                arySearchRt = Searching_Douban(strQueryTxt);
                                                          
                break;
                
            case "mtime":
                
                arySearchRt = Searching_Mtime(strQueryTxt);
                
                break;
                
        }
    
		WScript.Echo(CONST_LIST_LEVEL_1 + "找到影片数目：" + arySearchRt.length);
		WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);
        
        if (arySearchRt.length > 0) {
            blnGetAgain = false;
        }
    } 

	for (var i = 0; i < arySearchRt.length; i++) {
		WScript.Echo(CONST_LIST_LEVEL_1 + "(" + (i + 1) + ") " + arySearchRt[i][1]);
	}

	WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);

	var blnSelected = false;

	while (blnSelected == false) {
		strMovieIdx = AskForOption(CONST_LIST_LEVEL_1 + "请选择符合的影片(0=跳过；r=重新搜索)：");
                   
		if (strMovieIdx == "0") {
			WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);
            
            m_strFileOption = strMovieIdx;

			return "";
		}

		if (strMovieIdx == "r") {
			WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);

			strQueryTxt = AskForOption(CONST_LIST_LEVEL_1 + "请输入新的关键字重新搜索影片：");

			return DisplaySearching(strSearchIn, strQueryTxt);
		}
        
        var intMovieIdx = parseInt(strMovieIdx);

		if (intMovieIdx > 0 && intMovieIdx <= arySearchRt.length) {
			return arySearchRt[intMovieIdx - 1][0];
		}
	}

	return "";
}


function Searching_Douban(strQueryTxt) {
	var strQueryKey = "0a0a604697c5185a1e1f20d3c74f490e";
	var strQueryUrl = "https://api.douban.com/v2/movie/search?q=";
    var arySearchAt = new Array();

	m_objReq.open("GET", strQueryUrl + encodeURIComponent(strQueryTxt), false);
	m_objReq.send();

	try {
		var colMsg = eval("(" + m_objReq.responseText + ")");

		if (colMsg.code) {
			WScript.Echo(CONST_LIST_LEVEL_1 + "Error of Douban: ");
            WScript.Echo(CONST_LIST_LEVEL_2 + "code: " + colMsg.code);
			WScript.Echo(CONST_LIST_LEVEL_2 + "msg:  " + colMsg.msg);
			WScript.Echo(CONST_LIST_LEVEL_2 + "request: " + colMsg.request);

			return arySearchAt;
		}

		for (var i = 0; i < colMsg.subjects.length; i++) {
			var strEchoText = "";

			strEchoText += colMsg.subjects[i].title;
			strEchoText += " [";
			strEchoText += colMsg.subjects[i].year + ", ";
			strEchoText += colMsg.subjects[i].genres.join("/") + ", ";
            strEchoText += colMsg.subjects[i].id;
			strEchoText += "]";

			arySearchAt[i] = new Array(colMsg.subjects[i].id, strEchoText);
		}

	} catch (e) {}

	return arySearchAt;
}


function Searching_Mtime(strQueryTxt) {
	var strQueryUrl = "http://m.mtime.cn/Service/callback.mi/Showtime/SearchVoice.api?keyword=";
    var arySearchAt = new Array();

	m_objReq.open("GET", strQueryUrl + encodeURIComponent(strQueryTxt), false);
	m_objReq.send();

	try {
		var colMsg = eval("(" + m_objReq.responseText + ")");

		for (var i = 0; i < colMsg.movies.length; i++) {
			var strEchoText = "";

			strEchoText += colMsg.movies[i].name;
			strEchoText += " [";
			strEchoText += colMsg.movies[i].year + ", ";
			strEchoText += colMsg.movies[i].movieType.replace(/ \| /g, "/") + ", ";
            strEchoText += colMsg.movies[i].id;
			strEchoText += "]";

			arySearchAt[i] = new Array(colMsg.movies[i].id, strEchoText);
		}

	} catch (e) {
        return arySearchAt;    
    }

	return arySearchAt;
}


function CreateNof_TVEpisodes(strFilePath, strFileName, strOption) {
    var strExtension = GetExtension(strFileName);
    
    WScript.Echo(CONST_SEPARATE_LINE_1);
        
    DisplayResult(CONST_LIST_LEVEL_1 + "发现影片文件：" + strFileName, "", false, " ");
    
    if (strOption != "a" && strOption != "s") {
        DisplayResult(CONST_LIST_LEVEL_1 + "生成剧集：*.nfo", CONST_RESULT_IGNORE, false);
        
        return false;
    }
    
    if (strOption == "s") {
        WScript.Echo(CONST_LIST_LEVEL_1 + CONST_SEPARATE_SHORT);
        
        strOption = AskForOption(CONST_LIST_LEVEL_1 + "NFO文件已存在，是否替换？(y=替换)：");
        
        if (strOption != "y") {
            DisplayResult(CONST_LIST_LEVEL_1 + "生成剧集：*.nfo", CONST_RESULT_IGNORE, false);
            
            return false;
        }
    }
            
    m_objXml.loadXML("<?xml version=\"1.0\" encoding=\"utf-8\" ?><tvshow />");
        
	var eleMovie = m_objXml.documentElement;
    
    var objNewNode = Xml_AddNode(m_objXml, eleMovie, "episodedetails", "");
      
    var strSeason = GetTVEpisodes_Season(strFilePath + strFileName);
    var strEpisode = GetTVEpisodes_Episode(strFilePath + strFileName);
    
    var s = "00" + strSeason;
    var e = "00" + strEpisode;
    
    var aryTitle = GetMovieKeyword(strFileName).split(" ");
    var strTitle = aryTitle.join(".");
    
    strTitle = strTitle.replace(/.s(eason)?.[0-9]{1,2}/gi, "");
    
    if (strSeason != "0") {
        strTitle = strTitle + ".S" + s.substr(s.length - 2, 2) + "E" + e.substr(e.length - 2, 2);    
    } else {
        strTitle = strTitle + ".EP" + e.substr(e.length - 2, 2);
    }
     
	Xml_AddNode(m_objXml, objNewNode, "title", strTitle);
    
	Xml_AddNode(m_objXml, objNewNode, "season", strSeason);
	Xml_AddNode(m_objXml, objNewNode, "episode", strEpisode); 
    
    var strBookmark = "0";
    
    if (WScript.Arguments.length > 3) {
        var strBookmark = WScript.Arguments.Item(3);
    } 
    
    Xml_AddNode(m_objXml, objNewNode, "epbookmark", strBookmark); 		
   
    var strExtension = GetExtension(strFileName);
    
    m_objXml.save(strFilePath + strFileName.replace("." + strExtension, ".nfo"));
    
    DisplayResult(CONST_LIST_LEVEL_1 + "生成剧集：*.nfo", CONST_RESULT_OK, false);
    
    return true;
}


function GetMovieKeyword(strFileName) {
	var d = new Date();
    var strBreak = "|bdrip|webrip|blu-ray|rip|720p|1080p|bd720p|bd1080p|hd720p|hd1080p|tc720p|tc1080p|hr-hdtv|x264|aac|ac3|";
	var aryField = strFileName.replace(/\]\[/, ".").replace(/\[/, "").replace(/\]/, "").replace(/e(p)?[0-9]{1,2}/gi, "").split(".");
	var strKeyword = "";
    
    var expPattern = new RegExp("s(eason)?[0-9]{1,2}", "gi");    
    
	for (var i=0; i<aryField.length; i++) {
		try {
			var strField = aryField[i].toLowerCase();
            
            if (strBreak.indexOf("|" + strField + "|") > -1) {
                break;
            }
            
            var aryPattern = expPattern.exec(strField);
            
            if (aryPattern instanceof Array) {
            
                strField = aryPattern[0].replace(/s(eason)?[0]{0,1}/gi, "");
                    
                strKeyword += " Season " + strField;    
                
                break;
            }

			var intYears = parseInt(strField);

			if ((intYears > 1950) && (intYears <= d.getYear())) {
				break;
			} else {
				strKeyword += " " + aryField[i];
			}
		} catch (e) {
            WScript.Echo(e.toString());
            
            strKeyword = aryField.join(" ");
        }

	}

	return strKeyword.replace(/(^\s*)|(\s*$)/g, "");
}


function Xml_AddNode(xmlDom, Parent, Node_Name, Node_Text) {
	var objNode = xmlDom.createElement(Node_Name);
    
	objNode.text = Node_Text;
	Parent.appendChild(objNode);
	return objNode;
}


function GetExtension(strFilename) {
	var aryParts = strFilename.split(".");
    
	return aryParts[aryParts.length - 1];
}


function GetMovie_Title(strFullPath, strDefault) {
    var aryFileName = strFullPath.split("\\");    
    var strFileName = aryFileName[aryFileName.length - 1];
    
    var expPattern = new RegExp("^\\[\\S*?]", "gi");
    var aryPattern = expPattern.exec(strFileName);
    
    if ((aryPattern instanceof Array) == false) {
        return strDefault;    
    }
    
    var strMovTitle = aryPattern[0];
    
    return strMovTitle.replace("[", "").replace("]", "");
}


function GetMovieSet_Set(strFullPath) {
    var aryFolder = strFullPath.split("\\");
    
    if (aryFolder.length < 3) {
        return "UnKnown";
    }
    
    return aryFolder[aryFolder.length - 3].replace("[", "").replace("]", "");
}


function GetMovieSet_Sorttitle(strFullPath) {
    var aryFolder = strFullPath.split("\\");
    
    if (aryFolder.length < 2) {
        return aryFolder[aryFolder.length - 1];
    }

    var strFolder = aryFolder[aryFolder.length - 2];
    
    aryFolder = strFolder.split("][");
    
    return aryFolder[1];
}


function GetTVEpisodes_Season(strFullPath) {
    var aryFieldGet = strFullPath.split("\\");
    var strFileName = aryFieldGet[aryFieldGet.length - 1];
    
    var expPattern = new RegExp("s(eason)?[0-9]{1,2}", "gi");
    var arySeasons = expPattern.exec(strFileName);
    
    if ((arySeasons instanceof Array) == false) {
        return "0";    
    }
    
    strSeasons = arySeasons[0].replace(/s(eason)?/gi, "");
    
    if (strSeasons == "08") { strSeasons = "8"; }
    if (strSeasons == "09") { strSeasons = "9"; }
        
    return parseInt(strSeasons);
}


function GetTVEpisodes_Episode(strFullPath) {
    var aryFieldGet = strFullPath.split("\\");
    var strFileName = aryFieldGet[aryFieldGet.length - 1];
    
    var expPattern = new RegExp("e(p)?[0-9]{1,2}", "gi");
    var aryEpisode = expPattern.exec(strFileName);
    
    if ((aryEpisode instanceof Array) == false) {
        return "1";
    }
    
    strEpisode = aryEpisode[0].replace(/e(p)?/gi, "");
    
    if (strEpisode == "08") { strEpisode = "8"; }
    if (strEpisode == "09") { strEpisode = "9"; }
        
    return parseInt(strEpisode);
}


function GetTVEpisodes_Studio(strFullPath) {
    var STUDIO_DEFINED = new Array("BBC", "CCTV", "PBS", "NHK", "IMAX", "Discovery", "CBS", "Fox", "国家地理");
    
    for (var i = 0; i < STUDIO_DEFINED.length; i++) {
        if (strFullPath.lastIndexOf(STUDIO_DEFINED[i]) > -1) {
            return STUDIO_DEFINED[i];
        }
    }

    return "UnKnown";
}


//********************************************************************
//* Function: FormatDateTime
//*
//* Purpose: Return the time format in YYYY-MM-DDTHH:mm:ss.
//*
//* Input:
//*  [in]    lngTime	The value of time, 0 for now.
//* Output:
//*  [out]	 string		Formatted time.
//*
//********************************************************************
function FormatDate(strTime) {
	var myDate;
	var strFmt;

	myDate=new Date(strTime);

	strFmt  = myDate.getYear() + ".";
	strFmt += StringAlign(myDate.getMonth() + 1, -2, "0") + ".";
	strFmt += StringAlign(myDate.getDate(), -2, "0");

	return strFmt;
}


//********************************************************************
//* Function: GetCharCount
//*
//* Purpose: Calculate the chars of a string.
//*
//* Input:
//*  [in]    strString		The string.
//* Output:
//*  [out]	 integer		The count of chars.
//*
//********************************************************************
function GetCharCount(strString) {
	var j = 0;
	var k = 0;

	strString = strString.toString();

	k = strString.length;

	for (i = 0; i < k; i++) {
		j++;

		if (strString.substr(i, 1).charCodeAt() > 256) {
			j++;
		}
	}

	return j;
}


//********************************************************************
//* Function: StringPadding
//*
//* Purpose: Padding a string.
//*
//* Input:
//*  [in]    strString		Input string.
//*  [in]    intLength		Length of string after padding, positive or negative.
//*  [in]    strPadding 	Padding string.
//* Output:
//*  [out]	 string 		Padding result.
//*
//********************************************************************
function StringPadding(strString, intLength, strPadding) {
	var blnNegative = false;
	var strAssemble = "";
	var intCharCnts = 0;

	if (intLength < 0) {
		blnNegative = true;
	}

	intLength = Math.abs(intLength);
	intCharCnts = GetCharCount(strString);

	if (intLength < intCharCnts) {
		return strString;
	}

	while (!(strAssemble.length > (intLength - intCharCnts))) {
		strAssemble += strPadding;
	}

	strAssemble = strAssemble.substring(0, intLength - intCharCnts);

	if (blnNegative == false) {
		return strString + strAssemble;
	} else {
		return strAssemble + strString;
	}
}


//********************************************************************
//* Function: StringAlign
//*
//* Purpose: Output string in left, center, right.
//*
//* Input:
//*  [in]    strString		Input string.
//*  [in]    intLength		Length of string after padding, positive or negative.
//*  [in]    strPadding 	Padding string.
//* Output:
//*  [out]	 string 		Aligned string.
//*
//********************************************************************
function StringAlign(strString, intLength, strPadding) {
	var blnNegative = false;
	var strAssemble = "";
	var intCharCnts = 0;

	if (intLength > 0) {
		blnNegative = false;
	} else {
		blnNegative = true;
	}

	intLength = Math.abs(intLength);

	if (intLength <= strString.length) {
		return strString;
	}

	while (strAssemble.length < intLength) {
		strAssemble += strPadding;
	}

	intCharCnts = GetCharCount(strString);

	strAssemble = strAssemble.substring(0, intLength - intCharCnts);

	if (blnNegative == false) {
		return strString + strAssemble;
	} else {
		return strAssemble + strString;
	}
}


//********************************************************************
//* Function: DisplayResult
//*
//* Purpose: Display a result message in entire row.
//*
//* Input:
//*  [in]    strMessage		The message which will be displayed.
//*  [in]    strResult		The result which will be displayed.
//*  [in]    blnStdOut		True = WScript.Std.Out, False = WScript.Echo
//*  [in]    strSperator	The character of sperator.
//* Output:
//*  [out]	 none.
//*
//********************************************************************
function DisplayResult(strMessage, strResult, blnStdOut, strSperator) {
	var j = 0;
	var k = CONST_CHARS_IN_LINE - strResult.length;
	var m = 0;
	var n = GetCharCount(strMessage);
	var s = "";

	var strSperator = arguments[3] ? arguments[3] : ".";

	// 若输出的文本超过预留空间，则取输出文本的首尾进行输出
	if (n > k) {
		k = k - 19;

		s = "..." + strMessage.substr(strMessage.length - 16, 16);
	}

	strMessage = StringPadding(strMessage, CONST_CHARS_IN_LINE, strSperator);

	m = strMessage.length

	for (var i = 0; i < m; i++) {
		j++;

		if (strMessage.substr(i, 1).charCodeAt() > 256) {
			j++;
		}

		if (j > k) {
			strMessage = strMessage.substr(0, i);
			i = m;
		}
	}

	strMessage = strMessage + s + strResult;

	if (blnStdOut == true) {
		WScript.StdOut.Write(strMessage + "\r");
	} else {
		WScript.Echo(strMessage);
	}
}


//********************************************************************
//* Function: 	AskForOption
//*
//* Purpose: 	Display the options for choosen.
//*
//* Input:
//*  [in]    strMessage		tip of options.
//* Output:
//*  [out]	 integer 		value of options.
//*
//********************************************************************
function AskForOption(strMessage) {
	var strInput;

	strInput = "";
	WScript.StdOut.Write(strMessage);

	while(!(WScript.StdIn.AtEndOfLine)) {
	   	strInput +=  WScript.StdIn.Read(1);
	}

	WScript.StdIn.ReadLine();

	strInput = strInput.replace(/(^\s*)|(\s*$)/g, "");

	return(strInput);
}


/**
 * Returns internal [[Class]] property of an object
 *
 * Ecma-262, 15.2.4.2
 * Object.prototype.toString( )
 *
 * When the toString method is called, the following steps are taken: 
 * 1. Get the [[Class]] property of this object. 
 * 2. Compute a string value by concatenating the three strings "[object ", Result (1), and "]". 
 * 3. Return Result (2).
 *
 * __getClass(5); // => "Number"
 * __getClass({}); // => "Object"
 * __getClass(/foo/); // => "RegExp"
 * __getClass(''); // => "String"
 * __getClass(true); // => "Boolean"
 * __getClass([]); // => "Array"
 * __getClass(undefined); // => "Window"
 * __getClass(Element); // => "Constructor"
 *
 */
function __getClass(object){
    return Object.prototype.toString.call(object).match(/^\[object\s(.*)\]$/)[1];
};

// 扩展一下，用于检测各种对象类型：
// var is ={
//     types : ["Array", "Boolean", "Date", "Number", "Object", "RegExp", "String", "Window", "HTMLDocument"]
// };
// for(var i = 0, c; c = is.types[i ++ ]; ){
//     is[c] = (function(type){
//         return function(obj){
//            return Object.prototype.toString.call(obj) == "[object " + type + "]";
//         }
//     )(c);
// }
// alert(is.Array([])); // true
// alert(is.Date(new Date)); // true
// alert(is.RegExp(/reg/ig)); // true
//-----------------------------------------------------------------------------
//                            End of the Script
//-----------------------------------------------------------------------------
