# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.8.5 (default, Sep  3 2020, 21:29:08) [MSC v.1916 64 bit (AMD64)]
# From type library '{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}'
# On Mon Jan 25 15:53:27 2021
'HwpObject 1.0 Type Library'
makepy_version = '0.5.01'
python_version = 0x30805f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}')
MajorVersion = 1
MinorVersion = 0
LibraryFlags = 8
LCID = 0x0

from win32com.client import DispatchBaseClass
class IHwpObject(DispatchBaseClass):
	'IHwpObject Interface'
	CLSID = IID('{5E6A8276-CF1C-42B8-BCED-319548B02AF6}')
	coclass_clsid = IID('{2291CF00-64A1-4877-A9B4-68CFE89612D6}')

	def ArcType(self, ArcType=defaultNamedNotOptArg):
		'method ArcType'
		return self._oleobj_.InvokeTypes(30076, LCID, 1, (3, 0), ((8, 1),),ArcType
			)

	def AutoNumType(self, autonum=defaultNamedNotOptArg):
		'method AutoNumType'
		return self._oleobj_.InvokeTypes(30057, LCID, 1, (3, 0), ((8, 1),),autonum
			)

	def BorderShape(self, BorderType=defaultNamedNotOptArg):
		'method BorderShape'
		return self._oleobj_.InvokeTypes(30047, LCID, 1, (3, 0), ((8, 1),),BorderType
			)

	def BreakWordLatin(self, BreakLatinWord=defaultNamedNotOptArg):
		'method BreakWordLatin'
		return self._oleobj_.InvokeTypes(30054, LCID, 1, (3, 0), ((8, 1),),BreakLatinWord
			)

	def BrushType(self, BrushType=defaultNamedNotOptArg):
		'method BrushType'
		return self._oleobj_.InvokeTypes(30079, LCID, 1, (3, 0), ((8, 1),),BrushType
			)

	def Canonical(self, Canonical=defaultNamedNotOptArg):
		'method Canonical'
		return self._oleobj_.InvokeTypes(30072, LCID, 1, (3, 0), ((8, 1),),Canonical
			)

	def CellApply(self, CellApply=defaultNamedNotOptArg):
		'method CellApply'
		return self._oleobj_.InvokeTypes(30062, LCID, 1, (3, 0), ((8, 1),),CellApply
			)

	def CharShadowType(self, ShadowType=defaultNamedNotOptArg):
		'method CharShadowType'
		return self._oleobj_.InvokeTypes(30022, LCID, 1, (3, 0), ((8, 1),),ShadowType
			)

	def CheckXObject(self, bstring=defaultNamedNotOptArg):
		'method CheckXObject'
		ret = self._oleobj_.InvokeTypes(30091, LCID, 1, (9, 0), ((8, 1),),bstring
			)
		if ret is not None:
			ret = Dispatch(ret, 'CheckXObject', None)
		return ret

	def Clear(self, option=defaultNamedNotOptArg):
		'method Clear'
		return self._oleobj_.InvokeTypes(10035, LCID, 1, (24, 0), ((12, 1),),option
			)

	def ColDefType(self, ColDefType=defaultNamedNotOptArg):
		'method ColDefType'
		return self._oleobj_.InvokeTypes(30031, LCID, 1, (3, 0), ((8, 1),),ColDefType
			)

	def ColLayoutType(self, ColLayoutType=defaultNamedNotOptArg):
		'method ColLayoutType'
		return self._oleobj_.InvokeTypes(30032, LCID, 1, (3, 0), ((8, 1),),ColLayoutType
			)

	def ConvertPUAHangulToUnicode(self, Text=defaultNamedNotOptArg):
		'method ConvertPUAHangulToUnicode'
		return self._oleobj_.InvokeTypes(30200, LCID, 1, (3, 0), ((12, 1),),Text
			)

	def CreateAction(self, actidstr=defaultNamedNotOptArg):
		'method CreateAction'
		ret = self._oleobj_.InvokeTypes(10031, LCID, 1, (9, 0), ((8, 1),),actidstr
			)
		if ret is not None:
			ret = Dispatch(ret, 'CreateAction', None)
		return ret

	def CreateField(self, Direction=defaultNamedNotOptArg, memo=defaultNamedNotOptArg, name=defaultNamedNotOptArg):
		'method CreateField'
		return self._oleobj_.InvokeTypes(10005, LCID, 1, (11, 0), ((8, 1), (12, 1), (12, 1)),Direction
			, memo, name)

	def CreateID(self, CreationID=defaultNamedNotOptArg):
		'method CreateID'
		return self._oleobj_.InvokeTypes(30035, LCID, 1, (3, 0), ((8, 1),),CreationID
			)

	def CreateMode(self, CreationMode=defaultNamedNotOptArg):
		'method CreateMode'
		return self._oleobj_.InvokeTypes(30034, LCID, 1, (3, 0), ((8, 1),),CreationMode
			)

	def CreatePageImage(self, Path=defaultNamedNotOptArg, pgno=defaultNamedNotOptArg, resolution=defaultNamedNotOptArg, depth=defaultNamedNotOptArg
			, Format=defaultNamedNotOptArg):
		'method CreatePageImage'
		return self._oleobj_.InvokeTypes(10025, LCID, 1, (11, 0), ((8, 1), (12, 1), (12, 1), (12, 1), (12, 1)),Path
			, pgno, resolution, depth, Format)

	def CreateSet(self, setidstr=defaultNamedNotOptArg):
		'method CreateSet'
		ret = self._oleobj_.InvokeTypes(30102, LCID, 1, (9, 0), ((8, 1),),setidstr
			)
		if ret is not None:
			ret = Dispatch(ret, 'CreateSet', None)
		return ret

	def CrookedSlash(self, CrookedSlash=defaultNamedNotOptArg):
		'method CrookedSlash'
		return self._oleobj_.InvokeTypes(30086, LCID, 1, (3, 0), ((8, 1),),CrookedSlash
			)

	def DSMark(self, DiacSymMark=defaultNamedNotOptArg):
		'method DSMark'
		return self._oleobj_.InvokeTypes(30024, LCID, 1, (3, 0), ((8, 1),),DiacSymMark
			)

	def DbfCodeType(self, DbfCode=defaultNamedNotOptArg):
		'method DbfCodeType'
		return self._oleobj_.InvokeTypes(30067, LCID, 1, (3, 0), ((8, 1),),DbfCode
			)

	def DeleteCtrl(self, ctrl=defaultNamedNotOptArg):
		'method DeleteCtrl'
		return self._oleobj_.InvokeTypes(10033, LCID, 1, (11, 0), ((9, 1),),ctrl
			)

	def Delimiter(self, Delimiter=defaultNamedNotOptArg):
		'method Delimiter'
		return self._oleobj_.InvokeTypes(30066, LCID, 1, (3, 0), ((8, 1),),Delimiter
			)

	def DrawAspect(self, DrawAspect=defaultNamedNotOptArg):
		'method DrawAspect'
		return self._oleobj_.InvokeTypes(30061, LCID, 1, (3, 0), ((8, 1),),DrawAspect
			)

	def DrawFillImage(self, fillimage=defaultNamedNotOptArg):
		'method DrawFillImage'
		return self._oleobj_.InvokeTypes(30014, LCID, 1, (3, 0), ((8, 1),),fillimage
			)

	def DrawShadowType(self, ShadowType=defaultNamedNotOptArg):
		'method DrawShadowType'
		return self._oleobj_.InvokeTypes(30015, LCID, 1, (3, 0), ((8, 1),),ShadowType
			)

	def Encrypt(self, Encrypt=defaultNamedNotOptArg):
		'method Encrypt'
		return self._oleobj_.InvokeTypes(30074, LCID, 1, (3, 0), ((8, 1),),Encrypt
			)

	def EndSize(self, EndSize=defaultNamedNotOptArg):
		'method EndSize'
		return self._oleobj_.InvokeTypes(30012, LCID, 1, (3, 0), ((8, 1),),EndSize
			)

	def EndStyle(self, EndStyle=defaultNamedNotOptArg):
		'method EndStyle'
		return self._oleobj_.InvokeTypes(30011, LCID, 1, (3, 0), ((8, 1),),EndStyle
			)

	def ExportStyle(self, param=defaultNamedNotOptArg):
		'method ExportStyle'
		return self._oleobj_.InvokeTypes(30001, LCID, 1, (11, 0), ((9, 1),),param
			)

	def FieldExist(self, Field=defaultNamedNotOptArg):
		'method FieldExist'
		return self._oleobj_.InvokeTypes(10007, LCID, 1, (11, 0), ((8, 1),),Field
			)

	def FillAreaType(self, FillArea=defaultNamedNotOptArg):
		'method FillAreaType'
		return self._oleobj_.InvokeTypes(30018, LCID, 1, (3, 0), ((8, 1),),FillArea
			)

	def FindCtrl(self):
		'method FindCtrl'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(30003, LCID, 1, (8, 0), (),)

	def FindDir(self, FindDir=defaultNamedNotOptArg):
		'method FindDir'
		return self._oleobj_.InvokeTypes(30068, LCID, 1, (3, 0), ((8, 1),),FindDir
			)

	def FindPrivateInfo(self, PrivateType=defaultNamedNotOptArg, PrivateString=defaultNamedOptArg):
		'method FindPrivateInfo'
		return self._oleobj_.InvokeTypes(30109, LCID, 1, (3, 0), ((3, 1), (12, 16)),PrivateType
			, PrivateString)

	def FontType(self, FontType=defaultNamedNotOptArg):
		'method FontType'
		return self._oleobj_.InvokeTypes(30019, LCID, 1, (3, 0), ((8, 1),),FontType
			)

	def GetBinDataPath(self, binid=defaultNamedNotOptArg):
		'method GetBinDataPath'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(30100, LCID, 1, (8, 0), ((3, 1),),binid
			)

	def GetCurFieldName(self, option=defaultNamedNotOptArg):
		'method GetCurFieldName'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(10011, LCID, 1, (8, 0), ((12, 1),),option
			)

	def GetFieldList(self, Number=defaultNamedNotOptArg, option=defaultNamedNotOptArg):
		'method GetFieldList'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(10015, LCID, 1, (8, 0), ((12, 1), (12, 1)),Number
			, option)

	def GetFieldText(self, Field=defaultNamedNotOptArg):
		'method GetFieldText'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(10008, LCID, 1, (8, 0), ((8, 1),),Field
			)

	def GetFileInfo(self, filename=defaultNamedNotOptArg):
		'method GetFileInfo'
		ret = self._oleobj_.InvokeTypes(30094, LCID, 1, (9, 0), ((8, 1),),filename
			)
		if ret is not None:
			ret = Dispatch(ret, 'GetFileInfo', None)
		return ret

	def GetFontList(self, langid=defaultNamedOptArg):
		'method GetFontList'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(30115, LCID, 1, (8, 0), ((12, 16),),langid
			)

	def GetHeadingString(self):
		'method GetHeadingString'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(30103, LCID, 1, (8, 0), (),)

	def GetMessageBoxMode(self):
		'method GetMessageBoxMode'
		return self._oleobj_.InvokeTypes(30098, LCID, 1, (3, 0), (),)

	def GetMousePos(self, XRelTo=defaultNamedNotOptArg, YRelTo=defaultNamedNotOptArg):
		'method GetMousePos'
		ret = self._oleobj_.InvokeTypes(10034, LCID, 1, (9, 0), ((3, 1), (3, 1)),XRelTo
			, YRelTo)
		if ret is not None:
			ret = Dispatch(ret, 'GetMousePos', None)
		return ret

	def GetPageText(self, pgno=defaultNamedNotOptArg, option=defaultNamedNotOptArg):
		'method GetPageText'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(30096, LCID, 1, (8, 0), ((3, 1), (12, 1)),pgno
			, option)

	def GetPos(self, List=pythoncom.Missing, Para=pythoncom.Missing, pos=pythoncom.Missing):
		'method GetPos'
		return self._ApplyTypes_(10020, 1, (24, 0), ((16387, 2), (16387, 2), (16387, 2)), 'GetPos', None,List
			, Para, pos)

	def GetPosBySet(self):
		'method GetPosBySet'
		ret = self._oleobj_.InvokeTypes(10040, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'GetPosBySet', None)
		return ret

	def GetScriptSource(self, filename=defaultNamedNotOptArg):
		'method GetScriptSource'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(30093, LCID, 1, (8, 0), ((8, 1),),filename
			)

	def GetSelectedPos(self, slist=pythoncom.Missing, spara=pythoncom.Missing, spos=pythoncom.Missing, elist=pythoncom.Missing
			, epara=pythoncom.Missing, epos=pythoncom.Missing):
		'method GetSelectedPos'
		return self._ApplyTypes_(30105, 1, (11, 0), ((16387, 2), (16387, 2), (16387, 2), (16387, 2), (16387, 2), (16387, 2)), 'GetSelectedPos', None,slist
			, spara, spos, elist, epara, epos
			)

	def GetSelectedPosBySet(self, sset=defaultNamedNotOptArg, eset=defaultNamedNotOptArg):
		'method GetSelectedPosBySet'
		return self._oleobj_.InvokeTypes(30106, LCID, 1, (11, 0), ((9, 1), (9, 1)),sset
			, eset)

	def GetText(self, Text=pythoncom.Missing):
		'method GetText'
		return self._ApplyTypes_(10019, 1, (3, 0), ((16392, 2),), 'GetText', None,Text
			)

	def GetTextFile(self, Format=defaultNamedNotOptArg, option=defaultNamedNotOptArg):
		'method GetTextFile'
		return self._ApplyTypes_(10023, 1, (12, 0), ((8, 1), (8, 1)), 'GetTextFile', None,Format
			, option)

	def GetUserInfo(self, userInfoId=defaultNamedNotOptArg):
		'method GetUserInfo'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(30120, LCID, 1, (8, 0), ((3, 1),),userInfoId
			)

	def Gradation(self, Gradation=defaultNamedNotOptArg):
		'method Gradation'
		return self._oleobj_.InvokeTypes(30016, LCID, 1, (3, 0), ((8, 1),),Gradation
			)

	def GridMethod(self, GridMethod=defaultNamedNotOptArg):
		'method GridMethod'
		return self._oleobj_.InvokeTypes(30059, LCID, 1, (3, 0), ((8, 1),),GridMethod
			)

	def GridViewLine(self, GridViewLine=defaultNamedNotOptArg):
		'method GridViewLine'
		return self._oleobj_.InvokeTypes(30060, LCID, 1, (3, 0), ((8, 1),),GridViewLine
			)

	def GutterMethod(self, GutterType=defaultNamedNotOptArg):
		'method GutterMethod'
		return self._oleobj_.InvokeTypes(30028, LCID, 1, (3, 0), ((8, 1),),GutterType
			)

	def HAlign(self, HAlign=defaultNamedNotOptArg):
		'method HAlign'
		return self._oleobj_.InvokeTypes(30036, LCID, 1, (3, 0), ((8, 1),),HAlign
			)

	def Handler(self, Handler=defaultNamedNotOptArg):
		'method Handler'
		return self._oleobj_.InvokeTypes(30033, LCID, 1, (3, 0), ((8, 1),),Handler
			)

	def Hash(self, Hash=defaultNamedNotOptArg):
		'method Hash'
		return self._oleobj_.InvokeTypes(30075, LCID, 1, (3, 0), ((8, 1),),Hash
			)

	def HatchStyle(self, HatchStyle=defaultNamedNotOptArg):
		'method HatchStyle'
		return self._oleobj_.InvokeTypes(30017, LCID, 1, (3, 0), ((8, 1),),HatchStyle
			)

	def HeadType(self, HeadingType=defaultNamedNotOptArg):
		'method HeadType'
		return self._oleobj_.InvokeTypes(30056, LCID, 1, (3, 0), ((8, 1),),HeadingType
			)

	def HeightRel(self, HeightRel=defaultNamedNotOptArg):
		'method HeightRel'
		return self._oleobj_.InvokeTypes(30042, LCID, 1, (3, 0), ((8, 1),),HeightRel
			)

	def Hiding(self, Hiding=defaultNamedNotOptArg):
		'method Hiding'
		return self._oleobj_.InvokeTypes(30080, LCID, 1, (3, 0), ((8, 1),),Hiding
			)

	def HorzRel(self, HorzRel=defaultNamedNotOptArg):
		'method HorzRel'
		return self._oleobj_.InvokeTypes(30040, LCID, 1, (3, 0), ((8, 1),),HorzRel
			)

	def HwpLineType(self, LineType=defaultNamedNotOptArg):
		'method HwpLineType'
		return self._oleobj_.InvokeTypes(30009, LCID, 1, (3, 0), ((8, 1),),LineType
			)

	def HwpLineWidth(self, LineWidth=defaultNamedNotOptArg):
		'method HwpLineWidth'
		return self._oleobj_.InvokeTypes(30008, LCID, 1, (3, 0), ((8, 1),),LineWidth
			)

	def HwpOutlineStyle(self, HwpOutlineStyle=defaultNamedNotOptArg):
		'method HwpOutlineStyle'
		return self._oleobj_.InvokeTypes(30013, LCID, 1, (3, 0), ((8, 1),),HwpOutlineStyle
			)

	def HwpOutlineType(self, HwpOutlineType=defaultNamedNotOptArg):
		'method HwpOutlineType'
		return self._oleobj_.InvokeTypes(30021, LCID, 1, (3, 0), ((8, 1),),HwpOutlineType
			)

	def HwpUnderlineShape(self, HwpUnderlineShape=defaultNamedNotOptArg):
		'method HwpUnderlineShape'
		return self._oleobj_.InvokeTypes(30090, LCID, 1, (3, 0), ((8, 1),),HwpUnderlineShape
			)

	def HwpUnderlineType(self, HwpUnderlineType=defaultNamedNotOptArg):
		'method HwpUnderlineType'
		return self._oleobj_.InvokeTypes(30020, LCID, 1, (3, 0), ((8, 1),),HwpUnderlineType
			)

	def HwpZoomType(self, ZoomType=defaultNamedNotOptArg):
		'method HwpZoomType'
		return self._oleobj_.InvokeTypes(30049, LCID, 1, (3, 0), ((8, 1),),ZoomType
			)

	def ImageFormat(self, ImageFormat=defaultNamedNotOptArg):
		'method ImageFormat'
		return self._oleobj_.InvokeTypes(30064, LCID, 1, (3, 0), ((8, 1),),ImageFormat
			)

	def ImportStyle(self, param=defaultNamedNotOptArg):
		'method ImportStyle'
		return self._oleobj_.InvokeTypes(30002, LCID, 1, (11, 0), ((9, 1),),param
			)

	def InitHParameterSet(self):
		'method InitHParameterSet'
		return self._oleobj_.InvokeTypes(10038, LCID, 1, (24, 0), (),)

	def InitScan(self, option=defaultNamedNotOptArg, Range=defaultNamedNotOptArg, spara=defaultNamedNotOptArg, spos=defaultNamedNotOptArg
			, epara=defaultNamedNotOptArg, epos=defaultNamedNotOptArg):
		'method InitScan'
		return self._oleobj_.InvokeTypes(10017, LCID, 1, (11, 0), ((12, 1), (12, 1), (12, 1), (12, 1), (12, 1), (12, 1)),option
			, Range, spara, spos, epara, epos
			)

	def Insert(self, Path=defaultNamedNotOptArg, Format=defaultNamedNotOptArg, arg=defaultNamedNotOptArg):
		'method Insert'
		return self._oleobj_.InvokeTypes(10003, LCID, 1, (24, 0), ((8, 1), (12, 1), (12, 1)),Path
			, Format, arg)

	def InsertBackgroundPicture(self, BorderType=defaultNamedNotOptArg, Path=defaultNamedNotOptArg, Embedded=defaultNamedNotOptArg, filloption=defaultNamedNotOptArg
			, watermark=defaultNamedNotOptArg, Effect=defaultNamedNotOptArg, Brightness=defaultNamedNotOptArg, Contrast=defaultNamedNotOptArg):
		'method InsertBackgroundPicture'
		return self._oleobj_.InvokeTypes(10030, LCID, 1, (11, 0), ((8, 1), (8, 1), (12, 1), (12, 1), (12, 1), (12, 1), (12, 1), (12, 1)),BorderType
			, Path, Embedded, filloption, watermark, Effect
			, Brightness, Contrast)

	def InsertCtrl(self, CtrlID=defaultNamedNotOptArg, initparam=defaultNamedOptArg):
		'method InsertCtrl'
		ret = self._oleobj_.InvokeTypes(10032, LCID, 1, (9, 0), ((8, 1), (12, 16)),CtrlID
			, initparam)
		if ret is not None:
			ret = Dispatch(ret, 'InsertCtrl', None)
		return ret

	def InsertPicture(self, Path=defaultNamedNotOptArg, Embedded=defaultNamedNotOptArg, sizeoption=defaultNamedNotOptArg, Reverse=defaultNamedNotOptArg
			, watermark=defaultNamedNotOptArg, Effect=defaultNamedNotOptArg, Width=defaultNamedNotOptArg, Height=defaultNamedNotOptArg):
		'method InsertPicture'
		ret = self._oleobj_.InvokeTypes(10029, LCID, 1, (9, 0), ((8, 1), (12, 1), (12, 1), (12, 1), (12, 1), (12, 1), (12, 1), (12, 1)),Path
			, Embedded, sizeoption, Reverse, watermark, Effect
			, Width, Height)
		if ret is not None:
			ret = Dispatch(ret, 'InsertPicture', None)
		return ret

	def IsActionEnable(self, actionID=defaultNamedNotOptArg):
		'method IsActionEnable'
		return self._oleobj_.InvokeTypes(30092, LCID, 1, (11, 0), ((8, 1),),actionID
			)

	def IsCommandLock(self, actionID=defaultNamedNotOptArg):
		'method IsCommandLock'
		return self._oleobj_.InvokeTypes(10028, LCID, 1, (11, 0), ((8, 1),),actionID
			)

	def KeyIndicator(self, seccnt=pythoncom.Missing, secno=pythoncom.Missing, prnpageno=pythoncom.Missing, colno=pythoncom.Missing
			, Line=pythoncom.Missing, pos=pythoncom.Missing, over=pythoncom.Missing, ctrlname=pythoncom.Missing):
		'method KeyIndicator'
		return self._ApplyTypes_(10022, 1, (11, 0), ((16387, 2), (16387, 2), (16387, 2), (16387, 2), (16387, 2), (16387, 2), (16386, 2), (16392, 2)), 'KeyIndicator', None,seccnt
			, secno, prnpageno, colno, Line, pos
			, over, ctrlname)

	def LineSpacingMethod(self, LineSpacing=defaultNamedNotOptArg):
		'method LineSpacingMethod'
		return self._oleobj_.InvokeTypes(30053, LCID, 1, (3, 0), ((8, 1),),LineSpacing
			)

	def LineWrapType(self, LineWrap=defaultNamedNotOptArg):
		'method LineWrapType'
		return self._oleobj_.InvokeTypes(30037, LCID, 1, (3, 0), ((8, 1),),LineWrap
			)

	def LockCommand(self, ActID=defaultNamedNotOptArg, isLock=defaultNamedNotOptArg):
		'method LockCommand'
		return self._oleobj_.InvokeTypes(10027, LCID, 1, (24, 0), ((8, 1), (11, 1)),ActID
			, isLock)

	def LunarToSolar(self, lYear=defaultNamedNotOptArg, lMonth=defaultNamedNotOptArg, lDay=defaultNamedNotOptArg, lLeap=defaultNamedNotOptArg
			, sYear=pythoncom.Missing, sMonth=pythoncom.Missing, sDay=pythoncom.Missing):
		'method LunarToSolar'
		return self._ApplyTypes_(30113, 1, (11, 0), ((3, 1), (3, 1), (3, 1), (11, 1), (16387, 2), (16387, 2), (16387, 2)), 'LunarToSolar', None,lYear
			, lMonth, lDay, lLeap, sYear, sMonth
			, sDay)

	def LunarToSolarBySet(self, lYear=defaultNamedNotOptArg, lMonth=defaultNamedNotOptArg, lDay=defaultNamedNotOptArg, lLeap=defaultNamedNotOptArg):
		'method LunarToSolarBySet'
		ret = self._oleobj_.InvokeTypes(30114, LCID, 1, (9, 0), ((3, 1), (3, 1), (3, 1), (11, 1)),lYear
			, lMonth, lDay, lLeap)
		if ret is not None:
			ret = Dispatch(ret, 'LunarToSolarBySet', None)
		return ret

	def MacroState(self, MacroState=defaultNamedNotOptArg):
		'method MacroState'
		return self._oleobj_.InvokeTypes(30081, LCID, 1, (3, 0), ((8, 1),),MacroState
			)

	def MailType(self, MailType=defaultNamedNotOptArg):
		'method MailType'
		return self._oleobj_.InvokeTypes(30065, LCID, 1, (3, 0), ((8, 1),),MailType
			)

	def MiliToHwpUnit(self, mili=defaultNamedNotOptArg):
		'method MiliToHwpUnit'
		return self._oleobj_.InvokeTypes(30005, LCID, 1, (3, 0), ((5, 1),),mili
			)

	def ModifyFieldProperties(self, Field=defaultNamedNotOptArg, remove=defaultNamedNotOptArg, Add=defaultNamedNotOptArg):
		'method ModifyFieldProperties'
		return self._oleobj_.InvokeTypes(10013, LCID, 1, (3, 0), ((8, 1), (3, 1), (3, 1)),Field
			, remove, Add)

	def MovePos(self, moveID=defaultNamedNotOptArg, Para=defaultNamedNotOptArg, pos=defaultNamedNotOptArg):
		'method MovePos'
		return self._oleobj_.InvokeTypes(10016, LCID, 1, (11, 0), ((12, 1), (12, 1), (12, 1)),moveID
			, Para, pos)

	def MoveToField(self, Field=defaultNamedNotOptArg, Text=defaultNamedNotOptArg, start=defaultNamedNotOptArg, select=defaultNamedNotOptArg):
		'method MoveToField'
		return self._oleobj_.InvokeTypes(10006, LCID, 1, (11, 0), ((8, 1), (12, 1), (12, 1), (12, 1)),Field
			, Text, start, select)

	def NumberFormat(self, NumFormat=defaultNamedNotOptArg):
		'method NumberFormat'
		return self._oleobj_.InvokeTypes(30026, LCID, 1, (3, 0), ((8, 1),),NumFormat
			)

	def Numbering(self, Numbering=defaultNamedNotOptArg):
		'method Numbering'
		return self._oleobj_.InvokeTypes(30077, LCID, 1, (3, 0), ((8, 1),),Numbering
			)

	def Open(self, filename=defaultNamedNotOptArg, Format=defaultNamedNotOptArg, arg=defaultNamedNotOptArg):
		'method Open'
		return self._oleobj_.InvokeTypes(10000, LCID, 1, (11, 0), ((8, 1), (12, 1), (12, 1)),filename
			, Format, arg)

	def PageNumPosition(self, pagenumpos=defaultNamedNotOptArg):
		'method PageNumPosition'
		return self._oleobj_.InvokeTypes(30058, LCID, 1, (3, 0), ((8, 1),),pagenumpos
			)

	def PageType(self, PageType=defaultNamedNotOptArg):
		'method PageType'
		return self._oleobj_.InvokeTypes(30030, LCID, 1, (3, 0), ((8, 1),),PageType
			)

	def ParaHeadAlign(self, ParaHeadAlign=defaultNamedNotOptArg):
		'method ParaHeadAlign'
		return self._oleobj_.InvokeTypes(30025, LCID, 1, (3, 0), ((8, 1),),ParaHeadAlign
			)

	def PicEffect(self, PicEffect=defaultNamedNotOptArg):
		'method PicEffect'
		return self._oleobj_.InvokeTypes(30010, LCID, 1, (3, 0), ((8, 1),),PicEffect
			)

	def PlacementType(self, Restart=defaultNamedNotOptArg):
		'method PlacementType'
		return self._oleobj_.InvokeTypes(30027, LCID, 1, (3, 0), ((8, 1),),Restart
			)

	def PointToHwpUnit(self, Point=defaultNamedNotOptArg):
		'method PointToHwpUnit'
		return self._oleobj_.InvokeTypes(30006, LCID, 1, (3, 0), ((5, 1),),Point
			)

	def PresentEffect(self, prsnteffect=defaultNamedNotOptArg):
		'method PresentEffect'
		return self._oleobj_.InvokeTypes(30089, LCID, 1, (3, 0), ((8, 1),),prsnteffect
			)

	def PrintDevice(self, PrintDevice=defaultNamedNotOptArg):
		'method PrintDevice'
		return self._oleobj_.InvokeTypes(30052, LCID, 1, (3, 0), ((8, 1),),PrintDevice
			)

	def PrintPaper(self, PrintPaper=defaultNamedNotOptArg):
		'method PrintPaper'
		return self._oleobj_.InvokeTypes(30078, LCID, 1, (3, 0), ((8, 1),),PrintPaper
			)

	def PrintRange(self, PrintRange=defaultNamedNotOptArg):
		'method PrintRange'
		return self._oleobj_.InvokeTypes(30050, LCID, 1, (3, 0), ((8, 1),),PrintRange
			)

	def PrintType(self, PrintMethod=defaultNamedNotOptArg):
		'method PrintType'
		return self._oleobj_.InvokeTypes(30051, LCID, 1, (3, 0), ((8, 1),),PrintMethod
			)

	def ProtectPrivateInfo(self, PotectingChar=defaultNamedNotOptArg, PrivatePatternType=defaultNamedOptArg):
		'method ProtectPrivateInfo'
		return self._oleobj_.InvokeTypes(30110, LCID, 1, (11, 0), ((8, 1), (12, 16)),PotectingChar
			, PrivatePatternType)

	def PutFieldText(self, Field=defaultNamedNotOptArg, Text=defaultNamedNotOptArg):
		'method PutFieldText'
		return self._oleobj_.InvokeTypes(10009, LCID, 1, (24, 0), ((8, 1), (8, 1)),Field
			, Text)

	def Quit(self):
		'method Quit'
		return self._oleobj_.InvokeTypes(30000, LCID, 1, (24, 0), (),)

	def RGBColor(self, red=defaultNamedNotOptArg, green=defaultNamedNotOptArg, blue=defaultNamedNotOptArg):
		'method RGBColor'
		return self._oleobj_.InvokeTypes(30007, LCID, 1, (3, 0), ((17, 1), (17, 1), (17, 1)),red
			, green, blue)

	def RegisterModule(self, ModuleType=defaultNamedNotOptArg, ModuleData=defaultNamedNotOptArg):
		'method RegisterModule'
		return self._oleobj_.InvokeTypes(10036, LCID, 1, (11, 0), ((8, 1), (12, 1)),ModuleType
			, ModuleData)

	def RegisterPrivateInfoPattern(self, PrivateType=defaultNamedNotOptArg, PrivatePattern=defaultNamedNotOptArg):
		'method RegisterPrivateInfoPattern'
		return self._oleobj_.InvokeTypes(30108, LCID, 1, (11, 0), ((3, 1), (8, 1)),PrivateType
			, PrivatePattern)

	def ReleaseAction(self, action=defaultNamedNotOptArg):
		'method ReleaseAction'
		return self._oleobj_.InvokeTypes(30118, LCID, 1, (24, 0), ((9, 1),),action
			)

	def ReleaseScan(self):
		'method ReleaseScan'
		return self._oleobj_.InvokeTypes(10018, LCID, 1, (24, 0), (),)

	def RenameField(self, oldname=defaultNamedNotOptArg, newname=defaultNamedNotOptArg):
		'method RenameField'
		return self._oleobj_.InvokeTypes(10010, LCID, 1, (24, 0), ((8, 1), (8, 1)),oldname
			, newname)

	def ReplaceAction(self, OldActionID=defaultNamedNotOptArg, NewActionID=defaultNamedNotOptArg):
		'method ReplaceAction'
		return self._oleobj_.InvokeTypes(10037, LCID, 1, (11, 0), ((8, 1), (8, 1)),OldActionID
			, NewActionID)

	def ReplaceFont(self, langid=defaultNamedNotOptArg, desFontName=defaultNamedNotOptArg, desFontType=defaultNamedNotOptArg, newFontName=defaultNamedNotOptArg
			, newFontType=defaultNamedNotOptArg):
		'method ReplaceFont'
		return self._oleobj_.InvokeTypes(30116, LCID, 1, (11, 0), ((3, 1), (8, 1), (3, 1), (8, 1), (3, 1)),langid
			, desFontName, desFontType, newFontName, newFontType)

	def Revision(self, Revision=defaultNamedNotOptArg):
		'method Revision'
		return self._oleobj_.InvokeTypes(30071, LCID, 1, (3, 0), ((8, 1),),Revision
			)

	def Run(self, ActID=defaultNamedNotOptArg):
		'method Run'
		return self._oleobj_.InvokeTypes(10026, LCID, 1, (24, 0), ((8, 1),),ActID
			)

	def RunScriptMacro(self, FunctionName=defaultNamedNotOptArg, uMacroType=defaultNamedNotOptArg, uScriptType=defaultNamedNotOptArg):
		'method RunScriptMacro'
		return self._oleobj_.InvokeTypes(30095, LCID, 1, (11, 0), ((8, 1), (3, 1), (3, 1)),FunctionName
			, uMacroType, uScriptType)

	def Save(self, save_if_dirty=defaultNamedNotOptArg):
		'method Save'
		return self._oleobj_.InvokeTypes(10001, LCID, 1, (11, 0), ((12, 1),),save_if_dirty
			)

	def SaveAs(self, Path=defaultNamedNotOptArg, Format=defaultNamedNotOptArg, arg=defaultNamedNotOptArg):
		'method SaveAs'
		return self._oleobj_.InvokeTypes(10002, LCID, 1, (11, 0), ((8, 1), (12, 1), (12, 1)),Path
			, Format, arg)

	def ScanFont(self):
		'method ScanFont'
		return self._oleobj_.InvokeTypes(30117, LCID, 1, (11, 0), (),)

	def SelectText(self, spara=defaultNamedNotOptArg, spos=defaultNamedNotOptArg, epara=defaultNamedNotOptArg, epos=defaultNamedNotOptArg):
		'method SelectText'
		return self._oleobj_.InvokeTypes(10004, LCID, 1, (11, 0), ((3, 1), (3, 1), (3, 1), (3, 1)),spara
			, spos, epara, epos)

	def SetBarCodeImage(self, lpImagePath=defaultNamedNotOptArg, pgno=defaultNamedNotOptArg, index=defaultNamedNotOptArg, X=defaultNamedNotOptArg
			, Y=defaultNamedNotOptArg, Width=defaultNamedNotOptArg, Height=defaultNamedNotOptArg):
		'method SetBarCodeImage'
		return self._oleobj_.InvokeTypes(30097, LCID, 1, (11, 0), ((8, 1), (3, 1), (3, 1), (3, 1), (3, 1), (3, 1), (3, 1)),lpImagePath
			, pgno, index, X, Y, Width
			, Height)

	def SetCurFieldName(self, Field=defaultNamedNotOptArg, option=defaultNamedNotOptArg, Direction=defaultNamedNotOptArg, memo=defaultNamedNotOptArg):
		'method SetCurFieldName'
		return self._oleobj_.InvokeTypes(10012, LCID, 1, (11, 0), ((8, 1), (12, 1), (8, 1), (8, 1)),Field
			, option, Direction, memo)

	def SetDRMAuthority(self, authority=defaultNamedNotOptArg):
		'method SetDRMAuthority'
		return self._oleobj_.InvokeTypes(30101, LCID, 1, (11, 0), ((3, 1),),authority
			)

	def SetFieldViewOption(self, option=defaultNamedNotOptArg):
		'method SetFieldViewOption'
		return self._oleobj_.InvokeTypes(10014, LCID, 1, (3, 0), ((3, 1),),option
			)

	def SetMessageBoxMode(self, Mode=defaultNamedNotOptArg):
		'method SetMessageBoxMode'
		return self._oleobj_.InvokeTypes(30099, LCID, 1, (3, 0), ((3, 1),),Mode
			)

	def SetPos(self, List=defaultNamedNotOptArg, Para=defaultNamedNotOptArg, pos=defaultNamedNotOptArg):
		'method SetPos'
		return self._oleobj_.InvokeTypes(10021, LCID, 1, (11, 0), ((3, 1), (3, 1), (3, 1)),List
			, Para, pos)

	def SetPosBySet(self, dispVal=defaultNamedNotOptArg):
		'method SetPosBySet'
		return self._oleobj_.InvokeTypes(10041, LCID, 1, (11, 0), ((9, 1),),dispVal
			)

	def SetPrivateInfoPassword(self, Password=defaultNamedNotOptArg):
		'method SetPrivateInfoPassword'
		return self._oleobj_.InvokeTypes(30107, LCID, 1, (11, 0), ((8, 1),),Password
			)

	def SetTextFile(self, data=defaultNamedNotOptArg, Format=defaultNamedNotOptArg, option=defaultNamedNotOptArg):
		'method SetTextFile'
		return self._oleobj_.InvokeTypes(10024, LCID, 1, (3, 0), ((12, 1), (8, 1), (8, 1)),data
			, Format, option)

	def SetTitleName(self, Title=defaultNamedNotOptArg):
		'method SetTitleName'
		return self._oleobj_.InvokeTypes(30104, LCID, 1, (24, 0), ((8, 1),),Title
			)

	def SetUserInfo(self, userInfoId=defaultNamedNotOptArg, Value=defaultNamedNotOptArg):
		'method SetUserInfo'
		return self._oleobj_.InvokeTypes(30119, LCID, 1, (24, 0), ((3, 1), (8, 1)),userInfoId
			, Value)

	def SideType(self, SideType=defaultNamedNotOptArg):
		'method SideType'
		return self._oleobj_.InvokeTypes(30046, LCID, 1, (3, 0), ((8, 1),),SideType
			)

	def Signature(self, Signature=defaultNamedNotOptArg):
		'method Signature'
		return self._oleobj_.InvokeTypes(30073, LCID, 1, (3, 0), ((8, 1),),Signature
			)

	def Slash(self, Slash=defaultNamedNotOptArg):
		'method Slash'
		return self._oleobj_.InvokeTypes(30085, LCID, 1, (3, 0), ((8, 1),),Slash
			)

	def SolarToLunar(self, sYear=defaultNamedNotOptArg, sMonth=defaultNamedNotOptArg, sDay=defaultNamedNotOptArg, lYear=pythoncom.Missing
			, lMonth=pythoncom.Missing, lDay=pythoncom.Missing, lLeap=pythoncom.Missing):
		'method SolarToLunar'
		return self._ApplyTypes_(30111, 1, (11, 0), ((3, 1), (3, 1), (3, 1), (16387, 2), (16387, 2), (16387, 2), (16395, 2)), 'SolarToLunar', None,sYear
			, sMonth, sDay, lYear, lMonth, lDay
			, lLeap)

	def SolarToLunarBySet(self, sYear=defaultNamedNotOptArg, sMonth=defaultNamedNotOptArg, sDay=defaultNamedNotOptArg):
		'method SolarToLunarBySet'
		ret = self._oleobj_.InvokeTypes(30112, LCID, 1, (9, 0), ((3, 1), (3, 1), (3, 1)),sYear
			, sMonth, sDay)
		if ret is not None:
			ret = Dispatch(ret, 'SolarToLunarBySet', None)
		return ret

	def SortDelimiter(self, SortDelimiter=defaultNamedNotOptArg):
		'method SortDelimiter'
		return self._oleobj_.InvokeTypes(30070, LCID, 1, (3, 0), ((8, 1),),SortDelimiter
			)

	def StrikeOut(self, StrikeOutType=defaultNamedNotOptArg):
		'method StrikeOut'
		return self._oleobj_.InvokeTypes(30023, LCID, 1, (3, 0), ((8, 1),),StrikeOutType
			)

	def StyleType(self, StyleType=defaultNamedNotOptArg):
		'method StyleType'
		return self._oleobj_.InvokeTypes(30088, LCID, 1, (3, 0), ((8, 1),),StyleType
			)

	def SubtPos(self, SubtPos=defaultNamedNotOptArg):
		'method SubtPos'
		return self._oleobj_.InvokeTypes(30063, LCID, 1, (3, 0), ((8, 1),),SubtPos
			)

	def TableBreak(self, PageBreak=defaultNamedNotOptArg):
		'method TableBreak'
		return self._oleobj_.InvokeTypes(30048, LCID, 1, (3, 0), ((8, 1),),PageBreak
			)

	def TableFormat(self, TableFormat=defaultNamedNotOptArg):
		'method TableFormat'
		return self._oleobj_.InvokeTypes(30082, LCID, 1, (3, 0), ((8, 1),),TableFormat
			)

	def TableSwapType(self, tableswap=defaultNamedNotOptArg):
		'method TableSwapType'
		return self._oleobj_.InvokeTypes(30069, LCID, 1, (3, 0), ((8, 1),),tableswap
			)

	def TableTarget(self, TableTarget=defaultNamedNotOptArg):
		'method TableTarget'
		return self._oleobj_.InvokeTypes(30083, LCID, 1, (3, 0), ((8, 1),),TableTarget
			)

	def TextAlign(self, TextAlign=defaultNamedNotOptArg):
		'method TextAlign'
		return self._oleobj_.InvokeTypes(30055, LCID, 1, (3, 0), ((8, 1),),TextAlign
			)

	def TextArtAlign(self, TextArtAlign=defaultNamedNotOptArg):
		'method TextArtAlign'
		return self._oleobj_.InvokeTypes(30045, LCID, 1, (3, 0), ((8, 1),),TextArtAlign
			)

	def TextDir(self, TextDirection=defaultNamedNotOptArg):
		'method TextDir'
		return self._oleobj_.InvokeTypes(30029, LCID, 1, (3, 0), ((8, 1),),TextDirection
			)

	def TextFlowType(self, TextFlow=defaultNamedNotOptArg):
		'method TextFlowType'
		return self._oleobj_.InvokeTypes(30044, LCID, 1, (3, 0), ((8, 1),),TextFlow
			)

	def TextWrapType(self, TextWrap=defaultNamedNotOptArg):
		'method TextWrapType'
		return self._oleobj_.InvokeTypes(30043, LCID, 1, (3, 0), ((8, 1),),TextWrap
			)

	def UnSelectCtrl(self):
		'method UnSelectCtrl'
		return self._oleobj_.InvokeTypes(30004, LCID, 1, (24, 0), (),)

	def VAlign(self, VAlign=defaultNamedNotOptArg):
		'method VAlign'
		return self._oleobj_.InvokeTypes(30038, LCID, 1, (3, 0), ((8, 1),),VAlign
			)

	def VertRel(self, VertRel=defaultNamedNotOptArg):
		'method VertRel'
		return self._oleobj_.InvokeTypes(30039, LCID, 1, (3, 0), ((8, 1),),VertRel
			)

	def ViewFlag(self, ViewFlag=defaultNamedNotOptArg):
		'method ViewFlag'
		return self._oleobj_.InvokeTypes(30084, LCID, 1, (3, 0), ((8, 1),),ViewFlag
			)

	def WatermarkBrush(self, WatermarkBrush=defaultNamedNotOptArg):
		'method WatermarkBrush'
		return self._oleobj_.InvokeTypes(30087, LCID, 1, (3, 0), ((8, 1),),WatermarkBrush
			)

	def WidthRel(self, WidthRel=defaultNamedNotOptArg):
		'method WidthRel'
		return self._oleobj_.InvokeTypes(30041, LCID, 1, (3, 0), ((8, 1),),WidthRel
			)

	_prop_map_get_ = {
		"Application": (20000, 2, (9, 0), (), "Application", None),
		"CellShape": (7, 2, (9, 0), (), "CellShape", None),
		"CharShape": (8, 2, (9, 0), (), "CharShape", None),
		"CurFieldState": (5, 2, (3, 0), (), "CurFieldState", None),
		"CurSelectedCtrl": (11, 2, (9, 0), (), "CurSelectedCtrl", None),
		"EditMode": (3, 2, (3, 0), (), "EditMode", None),
		"EngineProperties": (20102, 2, (9, 0), (), "EngineProperties", None),
		"HAction": (20005, 2, (9, 0), (), "HAction", None),
		"HParameterSet": (20004, 2, (9, 0), (), "HParameterSet", None),
		"HeadCtrl": (9, 2, (9, 0), (), "HeadCtrl", None),
		"IsEmpty": (2, 2, (11, 0), (), "IsEmpty", None),
		"IsModified": (1, 2, (11, 0), (), "IsModified", None),
		"IsPrivateInfoProtected": (16, 2, (11, 0), (), "IsPrivateInfoProtected", None),
		"LastCtrl": (10, 2, (9, 0), (), "LastCtrl", None),
		"PageCount": (6, 2, (3, 0), (), "PageCount", None),
		"ParaShape": (12, 2, (9, 0), (), "ParaShape", None),
		"ParentCtrl": (13, 2, (9, 0), (), "ParentCtrl", None),
		"Path": (15, 2, (8, 0), (), "Path", None),
		"SelectionMode": (4, 2, (3, 0), (), "SelectionMode", None),
		"Version": (20101, 2, (8, 0), (), "Version", None),
		"ViewProperties": (14, 2, (9, 0), (), "ViewProperties", None),
		"XHwpDocuments": (20002, 2, (9, 0), (), "XHwpDocuments", None),
		"XHwpMessageBox": (20001, 2, (9, 0), (), "XHwpMessageBox", None),
		"XHwpODBC": (20006, 2, (9, 0), (), "XHwpODBC", None),
		"XHwpWindows": (20003, 2, (9, 0), (), "XHwpWindows", None),
	}
	_prop_map_put_ = {
		"CellShape": ((7, LCID, 4, 0),()),
		"CharShape": ((8, LCID, 4, 0),()),
		"EditMode": ((3, LCID, 4, 0),()),
		"EngineProperties": ((20102, LCID, 4, 0),()),
		"ParaShape": ((12, LCID, 4, 0),()),
		"ViewProperties": ((14, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{5E6A8276-CF1C-42B8-BCED-319548B02AF6}", IHwpObject )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.8.5 (default, Sep  3 2020, 21:29:08) [MSC v.1916 64 bit (AMD64)]
# From type library '{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}'
# On Mon Jan 25 15:53:27 2021
'HwpObject 1.0 Type Library'
makepy_version = '0.5.01'
python_version = 0x30805f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}')
MajorVersion = 1
MinorVersion = 0
LibraryFlags = 8
LCID = 0x0

IHwpObject_vtables_dispatch_ = 1
IHwpObject_vtables_ = [
	(( 'IsModified' , 'pVal' , ), 1, (1, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'IsEmpty' , 'pVal' , ), 2, (2, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'EditMode' , 'pVal' , ), 3, (3, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'EditMode' , 'pVal' , ), 3, (3, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'SelectionMode' , 'pVal' , ), 4, (4, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'CurFieldState' , 'pVal' , ), 5, (5, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'PageCount' , 'pVal' , ), 6, (6, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'CellShape' , 'pdispVal' , ), 7, (7, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'CellShape' , 'pdispVal' , ), 7, (7, (), [ (9, 1, None, None) , ], 1 , 4 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'CharShape' , 'pdispVal' , ), 8, (8, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'CharShape' , 'pdispVal' , ), 8, (8, (), [ (9, 1, None, None) , ], 1 , 4 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'HeadCtrl' , 'pdispVal' , ), 9, (9, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'LastCtrl' , 'pdispVal' , ), 10, (10, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'CurSelectedCtrl' , 'pdispVal' , ), 11, (11, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'ParaShape' , 'pdispVal' , ), 12, (12, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'ParaShape' , 'pdispVal' , ), 12, (12, (), [ (9, 1, None, None) , ], 1 , 4 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'ParentCtrl' , 'pdispVal' , ), 13, (13, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'ViewProperties' , 'pdispVal' , ), 14, (14, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'ViewProperties' , 'pdispVal' , ), 14, (14, (), [ (9, 1, None, None) , ], 1 , 4 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'Path' , 'pVal' , ), 15, (15, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'IsPrivateInfoProtected' , 'pVal' , ), 16, (16, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'Open' , 'filename' , 'Format' , 'arg' , 'pVal' , 
			 ), 10000, (10000, (), [ (8, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'Save' , 'save_if_dirty' , 'pVal' , ), 10001, (10001, (), [ (12, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'SaveAs' , 'Path' , 'Format' , 'arg' , 'pVal' , 
			 ), 10002, (10002, (), [ (8, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'Insert' , 'Path' , 'Format' , 'arg' , ), 10003, (10003, (), [ 
			 (8, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , ], 1 , 1 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'SelectText' , 'spara' , 'spos' , 'epara' , 'epos' , 
			 'pVal' , ), 10004, (10004, (), [ (3, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , 
			 (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'CreateField' , 'Direction' , 'memo' , 'name' , 'pVal' , 
			 ), 10005, (10005, (), [ (8, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'MoveToField' , 'Field' , 'Text' , 'start' , 'select' , 
			 'pVal' , ), 10006, (10006, (), [ (8, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , 
			 (12, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'FieldExist' , 'Field' , 'pVal' , ), 10007, (10007, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'GetFieldText' , 'Field' , 'pVal' , ), 10008, (10008, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'PutFieldText' , 'Field' , 'Text' , ), 10009, (10009, (), [ (8, 1, None, None) , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'RenameField' , 'oldname' , 'newname' , ), 10010, (10010, (), [ (8, 1, None, None) , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'GetCurFieldName' , 'option' , 'pVal' , ), 10011, (10011, (), [ (12, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'SetCurFieldName' , 'Field' , 'option' , 'Direction' , 'memo' , 
			 'pVal' , ), 10012, (10012, (), [ (8, 1, None, None) , (12, 1, None, None) , (8, 1, None, None) , 
			 (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'ModifyFieldProperties' , 'Field' , 'remove' , 'Add' , 'pVal' , 
			 ), 10013, (10013, (), [ (8, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'SetFieldViewOption' , 'option' , 'pVal' , ), 10014, (10014, (), [ (3, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'GetFieldList' , 'Number' , 'option' , 'pVal' , ), 10015, (10015, (), [ 
			 (12, 1, None, None) , (12, 1, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'MovePos' , 'moveID' , 'Para' , 'pos' , 'pVal' , 
			 ), 10016, (10016, (), [ (12, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'InitScan' , 'option' , 'Range' , 'spara' , 'spos' , 
			 'epara' , 'epos' , 'pVal' , ), 10017, (10017, (), [ (12, 1, None, None) , 
			 (12, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'ReleaseScan' , ), 10018, (10018, (), [ ], 1 , 1 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'GetText' , 'Text' , 'pVal' , ), 10019, (10019, (), [ (16392, 2, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'GetPos' , 'List' , 'Para' , 'pos' , ), 10020, (10020, (), [ 
			 (16387, 2, None, None) , (16387, 2, None, None) , (16387, 2, None, None) , ], 1 , 1 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'SetPos' , 'List' , 'Para' , 'pos' , 'pVal' , 
			 ), 10021, (10021, (), [ (3, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'KeyIndicator' , 'seccnt' , 'secno' , 'prnpageno' , 'colno' , 
			 'Line' , 'pos' , 'over' , 'ctrlname' , 'pVal' , 
			 ), 10022, (10022, (), [ (16387, 2, None, None) , (16387, 2, None, None) , (16387, 2, None, None) , (16387, 2, None, None) , 
			 (16387, 2, None, None) , (16387, 2, None, None) , (16386, 2, None, None) , (16392, 2, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'GetTextFile' , 'Format' , 'option' , 'pVal' , ), 10023, (10023, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 408 , (3, 0, None, None) , 0 , )),
	(( 'SetTextFile' , 'data' , 'Format' , 'option' , 'pVal' , 
			 ), 10024, (10024, (), [ (12, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'CreatePageImage' , 'Path' , 'pgno' , 'resolution' , 'depth' , 
			 'Format' , 'pVal' , ), 10025, (10025, (), [ (8, 1, None, None) , (12, 1, None, None) , 
			 (12, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'Run' , 'ActID' , ), 10026, (10026, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'LockCommand' , 'ActID' , 'isLock' , ), 10027, (10027, (), [ (8, 1, None, None) , 
			 (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'IsCommandLock' , 'actionID' , 'pVal' , ), 10028, (10028, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'InsertPicture' , 'Path' , 'Embedded' , 'sizeoption' , 'Reverse' , 
			 'watermark' , 'Effect' , 'Width' , 'Height' , 'pdispVal' , 
			 ), 10029, (10029, (), [ (8, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , 
			 (12, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'InsertBackgroundPicture' , 'BorderType' , 'Path' , 'Embedded' , 'filloption' , 
			 'watermark' , 'Effect' , 'Brightness' , 'Contrast' , 'pVal' , 
			 ), 10030, (10030, (), [ (8, 1, None, None) , (8, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , 
			 (12, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'CreateAction' , 'actidstr' , 'pdispVal' , ), 10031, (10031, (), [ (8, 1, None, None) , 
			 (16393, 10, None, None) , ], 1 , 1 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'InsertCtrl' , 'CtrlID' , 'initparam' , 'pdispVal' , ), 10032, (10032, (), [ 
			 (8, 1, None, None) , (12, 16, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 480 , (3, 0, None, None) , 0 , )),
	(( 'DeleteCtrl' , 'ctrl' , 'pVal' , ), 10033, (10033, (), [ (9, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'GetMousePos' , 'XRelTo' , 'YRelTo' , 'pdispVal' , ), 10034, (10034, (), [ 
			 (3, 1, None, None) , (3, 1, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'Clear' , 'option' , ), 10035, (10035, (), [ (12, 1, None, None) , ], 1 , 1 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( 'RegisterModule' , 'ModuleType' , 'ModuleData' , 'pVal' , ), 10036, (10036, (), [ 
			 (8, 1, None, None) , (12, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'ReplaceAction' , 'OldActionID' , 'NewActionID' , 'pVal' , ), 10037, (10037, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'InitHParameterSet' , ), 10038, (10038, (), [ ], 1 , 1 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'GetPosBySet' , 'pdispVal' , ), 10040, (10040, (), [ (16393, 10, None, None) , ], 1 , 1 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'SetPosBySet' , 'dispVal' , 'pVal' , ), 10041, (10041, (), [ (9, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'pdispVal' , ), 20000, (20000, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'XHwpMessageBox' , 'pdispVal' , ), 20001, (20001, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'XHwpDocuments' , 'pdispVal' , ), 20002, (20002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'XHwpWindows' , 'pdispVal' , ), 20003, (20003, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'HParameterSet' , 'pdispVal' , ), 20004, (20004, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 584 , (3, 0, None, None) , 0 , )),
	(( 'HAction' , 'pdispVal' , ), 20005, (20005, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'XHwpODBC' , 'pdispVal' , ), 20006, (20006, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 600 , (3, 0, None, None) , 0 , )),
	(( 'Version' , 'pVal' , ), 20101, (20101, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'EngineProperties' , 'pdispVal' , ), 20102, (20102, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'EngineProperties' , 'pdispVal' , ), 20102, (20102, (), [ (9, 1, None, None) , ], 1 , 4 , 4 , 0 , 624 , (3, 0, None, None) , 0 , )),
	(( 'Quit' , ), 30000, (30000, (), [ ], 1 , 1 , 4 , 0 , 632 , (3, 0, None, None) , 0 , )),
	(( 'ExportStyle' , 'param' , 'pVal' , ), 30001, (30001, (), [ (9, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 640 , (3, 0, None, None) , 0 , )),
	(( 'ImportStyle' , 'param' , 'pVal' , ), 30002, (30002, (), [ (9, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 648 , (3, 0, None, None) , 0 , )),
	(( 'FindCtrl' , 'pVal' , ), 30003, (30003, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 656 , (3, 0, None, None) , 0 , )),
	(( 'UnSelectCtrl' , ), 30004, (30004, (), [ ], 1 , 1 , 4 , 0 , 664 , (3, 0, None, None) , 0 , )),
	(( 'MiliToHwpUnit' , 'mili' , 'hwpunit' , ), 30005, (30005, (), [ (5, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 672 , (3, 0, None, None) , 0 , )),
	(( 'PointToHwpUnit' , 'Point' , 'hwpunit' , ), 30006, (30006, (), [ (5, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 680 , (3, 0, None, None) , 0 , )),
	(( 'RGBColor' , 'red' , 'green' , 'blue' , 'Color' , 
			 ), 30007, (30007, (), [ (17, 1, None, None) , (17, 1, None, None) , (17, 1, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 688 , (3, 0, None, None) , 0 , )),
	(( 'HwpLineWidth' , 'LineWidth' , 'index' , ), 30008, (30008, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 696 , (3, 0, None, None) , 0 , )),
	(( 'HwpLineType' , 'LineType' , 'index' , ), 30009, (30009, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 704 , (3, 0, None, None) , 0 , )),
	(( 'PicEffect' , 'PicEffect' , 'index' , ), 30010, (30010, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 712 , (3, 0, None, None) , 0 , )),
	(( 'EndStyle' , 'EndStyle' , 'index' , ), 30011, (30011, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 720 , (3, 0, None, None) , 0 , )),
	(( 'EndSize' , 'EndSize' , 'index' , ), 30012, (30012, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 728 , (3, 0, None, None) , 0 , )),
	(( 'HwpOutlineStyle' , 'HwpOutlineStyle' , 'index' , ), 30013, (30013, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 736 , (3, 0, None, None) , 0 , )),
	(( 'DrawFillImage' , 'fillimage' , 'index' , ), 30014, (30014, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 744 , (3, 0, None, None) , 0 , )),
	(( 'DrawShadowType' , 'ShadowType' , 'index' , ), 30015, (30015, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 752 , (3, 0, None, None) , 0 , )),
	(( 'Gradation' , 'Gradation' , 'index' , ), 30016, (30016, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 760 , (3, 0, None, None) , 0 , )),
	(( 'HatchStyle' , 'HatchStyle' , 'index' , ), 30017, (30017, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 768 , (3, 0, None, None) , 0 , )),
	(( 'FillAreaType' , 'FillArea' , 'index' , ), 30018, (30018, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 776 , (3, 0, None, None) , 0 , )),
	(( 'FontType' , 'FontType' , 'index' , ), 30019, (30019, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 784 , (3, 0, None, None) , 0 , )),
	(( 'HwpUnderlineType' , 'HwpUnderlineType' , 'index' , ), 30020, (30020, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 792 , (3, 0, None, None) , 0 , )),
	(( 'HwpOutlineType' , 'HwpOutlineType' , 'index' , ), 30021, (30021, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 800 , (3, 0, None, None) , 0 , )),
	(( 'CharShadowType' , 'ShadowType' , 'index' , ), 30022, (30022, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 808 , (3, 0, None, None) , 0 , )),
	(( 'StrikeOut' , 'StrikeOutType' , 'index' , ), 30023, (30023, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 816 , (3, 0, None, None) , 0 , )),
	(( 'DSMark' , 'DiacSymMark' , 'index' , ), 30024, (30024, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 824 , (3, 0, None, None) , 0 , )),
	(( 'ParaHeadAlign' , 'ParaHeadAlign' , 'index' , ), 30025, (30025, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 832 , (3, 0, None, None) , 0 , )),
	(( 'NumberFormat' , 'NumFormat' , 'index' , ), 30026, (30026, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 840 , (3, 0, None, None) , 0 , )),
	(( 'PlacementType' , 'Restart' , 'index' , ), 30027, (30027, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 848 , (3, 0, None, None) , 0 , )),
	(( 'GutterMethod' , 'GutterType' , 'index' , ), 30028, (30028, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 856 , (3, 0, None, None) , 0 , )),
	(( 'TextDir' , 'TextDirection' , 'index' , ), 30029, (30029, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 864 , (3, 0, None, None) , 0 , )),
	(( 'PageType' , 'PageType' , 'index' , ), 30030, (30030, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 872 , (3, 0, None, None) , 0 , )),
	(( 'ColDefType' , 'ColDefType' , 'index' , ), 30031, (30031, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 880 , (3, 0, None, None) , 0 , )),
	(( 'ColLayoutType' , 'ColLayoutType' , 'index' , ), 30032, (30032, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 888 , (3, 0, None, None) , 0 , )),
	(( 'Handler' , 'Handler' , 'index' , ), 30033, (30033, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 896 , (3, 0, None, None) , 0 , )),
	(( 'CreateMode' , 'CreationMode' , 'index' , ), 30034, (30034, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 904 , (3, 0, None, None) , 0 , )),
	(( 'CreateID' , 'CreationID' , 'index' , ), 30035, (30035, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 912 , (3, 0, None, None) , 0 , )),
	(( 'HAlign' , 'HAlign' , 'index' , ), 30036, (30036, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 920 , (3, 0, None, None) , 0 , )),
	(( 'LineWrapType' , 'LineWrap' , 'index' , ), 30037, (30037, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 928 , (3, 0, None, None) , 0 , )),
	(( 'VAlign' , 'VAlign' , 'index' , ), 30038, (30038, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 936 , (3, 0, None, None) , 0 , )),
	(( 'VertRel' , 'VertRel' , 'index' , ), 30039, (30039, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 944 , (3, 0, None, None) , 0 , )),
	(( 'HorzRel' , 'HorzRel' , 'index' , ), 30040, (30040, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 952 , (3, 0, None, None) , 0 , )),
	(( 'WidthRel' , 'WidthRel' , 'index' , ), 30041, (30041, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 960 , (3, 0, None, None) , 0 , )),
	(( 'HeightRel' , 'HeightRel' , 'index' , ), 30042, (30042, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 968 , (3, 0, None, None) , 0 , )),
	(( 'TextWrapType' , 'TextWrap' , 'index' , ), 30043, (30043, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 976 , (3, 0, None, None) , 0 , )),
	(( 'TextFlowType' , 'TextFlow' , 'index' , ), 30044, (30044, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 984 , (3, 0, None, None) , 0 , )),
	(( 'TextArtAlign' , 'TextArtAlign' , 'index' , ), 30045, (30045, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 992 , (3, 0, None, None) , 0 , )),
	(( 'SideType' , 'SideType' , 'index' , ), 30046, (30046, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1000 , (3, 0, None, None) , 0 , )),
	(( 'BorderShape' , 'BorderType' , 'index' , ), 30047, (30047, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1008 , (3, 0, None, None) , 0 , )),
	(( 'TableBreak' , 'PageBreak' , 'index' , ), 30048, (30048, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1016 , (3, 0, None, None) , 0 , )),
	(( 'HwpZoomType' , 'ZoomType' , 'index' , ), 30049, (30049, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1024 , (3, 0, None, None) , 0 , )),
	(( 'PrintRange' , 'PrintRange' , 'index' , ), 30050, (30050, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1032 , (3, 0, None, None) , 0 , )),
	(( 'PrintType' , 'PrintMethod' , 'index' , ), 30051, (30051, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1040 , (3, 0, None, None) , 0 , )),
	(( 'PrintDevice' , 'PrintDevice' , 'index' , ), 30052, (30052, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1048 , (3, 0, None, None) , 0 , )),
	(( 'LineSpacingMethod' , 'LineSpacing' , 'index' , ), 30053, (30053, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1056 , (3, 0, None, None) , 0 , )),
	(( 'BreakWordLatin' , 'BreakLatinWord' , 'index' , ), 30054, (30054, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1064 , (3, 0, None, None) , 0 , )),
	(( 'TextAlign' , 'TextAlign' , 'index' , ), 30055, (30055, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1072 , (3, 0, None, None) , 0 , )),
	(( 'HeadType' , 'HeadingType' , 'index' , ), 30056, (30056, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1080 , (3, 0, None, None) , 0 , )),
	(( 'AutoNumType' , 'autonum' , 'index' , ), 30057, (30057, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1088 , (3, 0, None, None) , 0 , )),
	(( 'PageNumPosition' , 'pagenumpos' , 'index' , ), 30058, (30058, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1096 , (3, 0, None, None) , 0 , )),
	(( 'GridMethod' , 'GridMethod' , 'index' , ), 30059, (30059, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1104 , (3, 0, None, None) , 0 , )),
	(( 'GridViewLine' , 'GridViewLine' , 'index' , ), 30060, (30060, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1112 , (3, 0, None, None) , 0 , )),
	(( 'DrawAspect' , 'DrawAspect' , 'index' , ), 30061, (30061, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1120 , (3, 0, None, None) , 0 , )),
	(( 'CellApply' , 'CellApply' , 'index' , ), 30062, (30062, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1128 , (3, 0, None, None) , 0 , )),
	(( 'SubtPos' , 'SubtPos' , 'index' , ), 30063, (30063, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1136 , (3, 0, None, None) , 0 , )),
	(( 'ImageFormat' , 'ImageFormat' , 'index' , ), 30064, (30064, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1144 , (3, 0, None, None) , 0 , )),
	(( 'MailType' , 'MailType' , 'index' , ), 30065, (30065, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1152 , (3, 0, None, None) , 0 , )),
	(( 'Delimiter' , 'Delimiter' , 'index' , ), 30066, (30066, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1160 , (3, 0, None, None) , 0 , )),
	(( 'DbfCodeType' , 'DbfCode' , 'index' , ), 30067, (30067, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1168 , (3, 0, None, None) , 0 , )),
	(( 'FindDir' , 'FindDir' , 'index' , ), 30068, (30068, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1176 , (3, 0, None, None) , 0 , )),
	(( 'TableSwapType' , 'tableswap' , 'index' , ), 30069, (30069, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1184 , (3, 0, None, None) , 0 , )),
	(( 'SortDelimiter' , 'SortDelimiter' , 'index' , ), 30070, (30070, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1192 , (3, 0, None, None) , 0 , )),
	(( 'Revision' , 'Revision' , 'index' , ), 30071, (30071, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1200 , (3, 0, None, None) , 0 , )),
	(( 'Canonical' , 'Canonical' , 'index' , ), 30072, (30072, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1208 , (3, 0, None, None) , 0 , )),
	(( 'Signature' , 'Signature' , 'index' , ), 30073, (30073, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1216 , (3, 0, None, None) , 0 , )),
	(( 'Encrypt' , 'Encrypt' , 'index' , ), 30074, (30074, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1224 , (3, 0, None, None) , 0 , )),
	(( 'Hash' , 'Hash' , 'index' , ), 30075, (30075, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1232 , (3, 0, None, None) , 0 , )),
	(( 'ArcType' , 'ArcType' , 'index' , ), 30076, (30076, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1240 , (3, 0, None, None) , 0 , )),
	(( 'Numbering' , 'Numbering' , 'index' , ), 30077, (30077, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1248 , (3, 0, None, None) , 0 , )),
	(( 'PrintPaper' , 'PrintPaper' , 'index' , ), 30078, (30078, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1256 , (3, 0, None, None) , 0 , )),
	(( 'BrushType' , 'BrushType' , 'index' , ), 30079, (30079, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1264 , (3, 0, None, None) , 0 , )),
	(( 'Hiding' , 'Hiding' , 'index' , ), 30080, (30080, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1272 , (3, 0, None, None) , 0 , )),
	(( 'MacroState' , 'MacroState' , 'index' , ), 30081, (30081, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1280 , (3, 0, None, None) , 0 , )),
	(( 'TableFormat' , 'TableFormat' , 'index' , ), 30082, (30082, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1288 , (3, 0, None, None) , 0 , )),
	(( 'TableTarget' , 'TableTarget' , 'index' , ), 30083, (30083, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1296 , (3, 0, None, None) , 0 , )),
	(( 'ViewFlag' , 'ViewFlag' , 'index' , ), 30084, (30084, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1304 , (3, 0, None, None) , 0 , )),
	(( 'Slash' , 'Slash' , 'index' , ), 30085, (30085, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1312 , (3, 0, None, None) , 0 , )),
	(( 'CrookedSlash' , 'CrookedSlash' , 'index' , ), 30086, (30086, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1320 , (3, 0, None, None) , 0 , )),
	(( 'WatermarkBrush' , 'WatermarkBrush' , 'index' , ), 30087, (30087, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1328 , (3, 0, None, None) , 0 , )),
	(( 'StyleType' , 'StyleType' , 'index' , ), 30088, (30088, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1336 , (3, 0, None, None) , 0 , )),
	(( 'PresentEffect' , 'prsnteffect' , 'index' , ), 30089, (30089, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1344 , (3, 0, None, None) , 0 , )),
	(( 'HwpUnderlineShape' , 'HwpUnderlineShape' , 'index' , ), 30090, (30090, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1352 , (3, 0, None, None) , 0 , )),
	(( 'CheckXObject' , 'bstring' , 'pdisp' , ), 30091, (30091, (), [ (8, 1, None, None) , 
			 (16393, 10, None, None) , ], 1 , 1 , 4 , 0 , 1360 , (3, 0, None, None) , 0 , )),
	(( 'IsActionEnable' , 'actionID' , 'pVal' , ), 30092, (30092, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1368 , (3, 0, None, None) , 0 , )),
	(( 'GetScriptSource' , 'filename' , 'pVal' , ), 30093, (30093, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1376 , (3, 0, None, None) , 0 , )),
	(( 'GetFileInfo' , 'filename' , 'pdispVal' , ), 30094, (30094, (), [ (8, 1, None, None) , 
			 (16393, 10, None, None) , ], 1 , 1 , 4 , 0 , 1384 , (3, 0, None, None) , 0 , )),
	(( 'RunScriptMacro' , 'FunctionName' , 'uMacroType' , 'uScriptType' , 'pVal' , 
			 ), 30095, (30095, (), [ (8, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1392 , (3, 0, None, None) , 0 , )),
	(( 'GetPageText' , 'pgno' , 'option' , 'pVal' , ), 30096, (30096, (), [ 
			 (3, 1, None, None) , (12, 1, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1400 , (3, 0, None, None) , 0 , )),
	(( 'SetBarCodeImage' , 'lpImagePath' , 'pgno' , 'index' , 'X' , 
			 'Y' , 'Width' , 'Height' , 'pVal' , ), 30097, (30097, (), [ 
			 (8, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , 
			 (3, 1, None, None) , (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1408 , (3, 0, None, None) , 0 , )),
	(( 'GetMessageBoxMode' , 'Mode' , ), 30098, (30098, (), [ (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1416 , (3, 0, None, None) , 0 , )),
	(( 'SetMessageBoxMode' , 'Mode' , 'oldmode' , ), 30099, (30099, (), [ (3, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1424 , (3, 0, None, None) , 0 , )),
	(( 'GetBinDataPath' , 'binid' , 'pVal' , ), 30100, (30100, (), [ (3, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1432 , (3, 0, None, None) , 0 , )),
	(( 'SetDRMAuthority' , 'authority' , 'pVal' , ), 30101, (30101, (), [ (3, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1440 , (3, 0, None, None) , 0 , )),
	(( 'CreateSet' , 'setidstr' , 'pdispVal' , ), 30102, (30102, (), [ (8, 1, None, None) , 
			 (16393, 10, None, None) , ], 1 , 1 , 4 , 0 , 1448 , (3, 0, None, None) , 0 , )),
	(( 'GetHeadingString' , 'pVal' , ), 30103, (30103, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1456 , (3, 0, None, None) , 0 , )),
	(( 'SetTitleName' , 'Title' , ), 30104, (30104, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1464 , (3, 0, None, None) , 0 , )),
	(( 'GetSelectedPos' , 'slist' , 'spara' , 'spos' , 'elist' , 
			 'epara' , 'epos' , 'pVal' , ), 30105, (30105, (), [ (16387, 2, None, None) , 
			 (16387, 2, None, None) , (16387, 2, None, None) , (16387, 2, None, None) , (16387, 2, None, None) , (16387, 2, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1472 , (3, 0, None, None) , 0 , )),
	(( 'GetSelectedPosBySet' , 'sset' , 'eset' , 'pVal' , ), 30106, (30106, (), [ 
			 (9, 1, None, None) , (9, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1480 , (3, 0, None, None) , 0 , )),
	(( 'SetPrivateInfoPassword' , 'Password' , 'pVal' , ), 30107, (30107, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1488 , (3, 0, None, None) , 0 , )),
	(( 'RegisterPrivateInfoPattern' , 'PrivateType' , 'PrivatePattern' , 'pVal' , ), 30108, (30108, (), [ 
			 (3, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1496 , (3, 0, None, None) , 0 , )),
	(( 'FindPrivateInfo' , 'PrivateType' , 'PrivateString' , 'pVal' , ), 30109, (30109, (), [ 
			 (3, 1, None, None) , (12, 16, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 1504 , (3, 0, None, None) , 0 , )),
	(( 'ProtectPrivateInfo' , 'PotectingChar' , 'PrivatePatternType' , 'pVal' , ), 30110, (30110, (), [ 
			 (8, 1, None, None) , (12, 16, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 1 , 1512 , (3, 0, None, None) , 0 , )),
	(( 'SolarToLunar' , 'sYear' , 'sMonth' , 'sDay' , 'lYear' , 
			 'lMonth' , 'lDay' , 'lLeap' , 'pVal' , ), 30111, (30111, (), [ 
			 (3, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , (16387, 2, None, None) , (16387, 2, None, None) , 
			 (16387, 2, None, None) , (16395, 2, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1520 , (3, 0, None, None) , 0 , )),
	(( 'SolarToLunarBySet' , 'sYear' , 'sMonth' , 'sDay' , 'pdispVal' , 
			 ), 30112, (30112, (), [ (3, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 0 , 1528 , (3, 0, None, None) , 0 , )),
	(( 'LunarToSolar' , 'lYear' , 'lMonth' , 'lDay' , 'lLeap' , 
			 'sYear' , 'sMonth' , 'sDay' , 'pVal' , ), 30113, (30113, (), [ 
			 (3, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , (11, 1, None, None) , (16387, 2, None, None) , 
			 (16387, 2, None, None) , (16387, 2, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1536 , (3, 0, None, None) , 0 , )),
	(( 'LunarToSolarBySet' , 'lYear' , 'lMonth' , 'lDay' , 'lLeap' , 
			 'pdispVal' , ), 30114, (30114, (), [ (3, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , 
			 (11, 1, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 0 , 1544 , (3, 0, None, None) , 0 , )),
	(( 'GetFontList' , 'langid' , 'pVal' , ), 30115, (30115, (), [ (12, 16, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 1 , 1552 , (3, 0, None, None) , 0 , )),
	(( 'ReplaceFont' , 'langid' , 'desFontName' , 'desFontType' , 'newFontName' , 
			 'newFontType' , 'pVal' , ), 30116, (30116, (), [ (3, 1, None, None) , (8, 1, None, None) , 
			 (3, 1, None, None) , (8, 1, None, None) , (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1560 , (3, 0, None, None) , 0 , )),
	(( 'ScanFont' , 'pVal' , ), 30117, (30117, (), [ (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1568 , (3, 0, None, None) , 0 , )),
	(( 'ReleaseAction' , 'action' , ), 30118, (30118, (), [ (9, 1, None, None) , ], 1 , 1 , 4 , 0 , 1576 , (3, 0, None, None) , 0 , )),
	(( 'SetUserInfo' , 'userInfoId' , 'Value' , ), 30119, (30119, (), [ (3, 1, None, None) , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1584 , (3, 0, None, None) , 0 , )),
	(( 'GetUserInfo' , 'userInfoId' , 'pVal' , ), 30120, (30120, (), [ (3, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1592 , (3, 0, None, None) , 0 , )),
	(( 'ConvertPUAHangulToUnicode' , 'Text' , 'pVal' , ), 30200, (30200, (), [ (12, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1600 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{5E6A8276-CF1C-42B8-BCED-319548B02AF6}", IHwpObject )
