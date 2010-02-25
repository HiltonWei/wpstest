{*************************************************************************************}
{                               Kingsoft Delphi Library
{
{
{                                           Create by Yeppy.Arch.WPS Date: 2004-3-22
{
{  Category:
{
{  Library Type: Runtime
{  Framework:    VCL / CLX
{  Platform:     Windows / Linux
{  Version:      1.0.0
{
{  Read first:
{
{
{  Documents:
{
{  To-do list:
{
{  Bug list:
{
{*************************************************************************************}

{//========================================
 #C1:2007-03-09:修正bug:26096

//=========================================}

unit DlgETFindReplace;

interface

uses
{$IFDEF VCL}
  Controls, ShellUtils, ShellImp,
{$ENDIF}
{$IFDEF CLX}
  QControls, QShellUtils, QShellImp,
{$ENDIF}
  Windows, Messages, SysUtils, Variants, Classes, StdCtrls, ComCtrls, Grip, FPSCtrls, MRUListCtrls,
  Shell_TLB, ET_TLB, et_tlbex, Forms, KSOUtils
  {TNT-USE}
  , TntComCtrls, TntStdCtrls, TntControls, GroupHeader
  {TNT-END};

type
  TFindReplaceMode = (frmFind, frmReplace, frmGoto);
  TDialogMode = (dmNormal, dmAdvanced);
  TSearchRange = (srSheet, srWorkBook);
  TSearchOrder = (soByRows, soByColumns);
  TSearchLookin = (slFormulas, slValues, slComments);
  TSearchOption = (soMatchCase, soMatchContents, soMatchByte);
  TSearchOptions = set of TSearchOption;

  TFindReplaceDlg = class(TShellDialog)
    tcFindReplace: TTntTabControl{TNT};
    mcbFind: TMRUComboBox;
    mcbReplace: TMRUComboBox;
    cbRange: TTntComboBox{TNT};
    cbOrder: TTntComboBox{TNT};
    cbLookin: TTntComboBox{TNT};
    cbMatchCase: TTntCheckBox{TNT};
    cbMatchContents: TTntCheckBox{TNT};
    cbMatchByte: TTntCheckBox{TNT};
    btnOptions: TTntButton{TNT};
    btnReplaceAll: TTntButton{TNT};
    btnReplace: TTntButton{TNT};
    btnFindAll: TTntButton{TNT};
    btnFindNext: TTntButton{TNT};
    btnClose: TTntButton{TNT};
    lblFind: TTntLabel{TNT};
    lblReplace: TTntLabel{TNT};
    lblWithin: TTntLabel{TNT};
    lblSearch: TTntLabel{TNT};
    lblLookin: TTntLabel{TNT};
    gpResizer: TGrip;
    tcGoto: TTntTabControl;
    btnGoto: TTntButton;
    ghSelect: TKSOGroupHeader;
    rbData: TTntRadioButton;
    rbComments: TTntRadioButton;
    rbBlanks: TTntRadioButton;
    rbLastCell: TTntRadioButton;
    rbVisibleCellsOnly: TTntRadioButton;
    cbConstants: TTntCheckBox;
    cbFormulas: TTntCheckBox;
    cbErrors: TTntCheckBox;
    cbText: TTntCheckBox;
    cbLogicals: TTntCheckBox;
    cbNumbers: TTntCheckBox;
    lblDataType: TTntLabel;
    rbHyperlink: TTntRadioButton;
    rbCurDataArea: TTntRadioButton;
    rbObject: TTntRadioButton;
    lvFindAll: TTntListView;
    procedure FormDeactivate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure tcFindReplaceChange(Sender: TObject);
    procedure tcFindReplaceEnter(Sender: TObject);
    procedure mcbFindChange(Sender: TObject);
    procedure mcbReplaceChange(Sender: TObject);
    procedure cbRangeSelect(Sender: TObject);
    procedure cbOrderSelect(Sender: TObject);
    procedure cbLookinSelect(Sender: TObject);
    procedure cbMatchCaseClick(Sender: TObject);
    procedure cbMatchContentsClick(Sender: TObject);
    procedure cbMatchByteClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnOptionsClick(Sender: TObject);
    procedure btnFindNextClick(Sender: TObject);
    procedure btnFindAllClick(Sender: TObject);
    procedure btnReplaceClick(Sender: TObject);
    procedure btnReplaceAllClick(Sender: TObject);
    procedure OnEnterSearch(Sender: TObject);
    procedure OnExitSearch(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure rbDataClick(Sender: TObject);
    procedure btnGotoClick(Sender: TObject);
    procedure cbNumbersClick(Sender: TObject);
    procedure cbConstantsClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure OnSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure OnColClick(Sender: TObject; Column: TListColumn);
    procedure OnClick(Sender: TObject);
    procedure lvFindAllResize(Sender: TObject);
    procedure lvFindAllCompare(Sender: TObject; Item1, Item2: TListItem;
      Data: Integer; var Compare: Integer);
  private
    { Private declarations }
    _FindNextFlag: Boolean;
    _ReplaceFirstFlag: Boolean; //第一次替换
    _ReplaceSecondFlag: Boolean; //第二次替换
    _ReplaceSinceFlag: Boolean; //以后的替换
    _ReplaceCompleteFlag: Boolean; //替换完成
    _FindCompleteFlag: Boolean; //查找完成
    _FirstFindInSheet: Boolean;
    FindFirst: IETRange;//查找全部的第一个

    SheetsCount: Integer; //工作簿中工作表总数

    //FRange: IETRange;
    FindReplaceRange: IETRange;
    FLocateRange: IETRange;

    FirstFindRangeRow: integer; //在某一工作表中找到的第一个单元格的行号，用于在适当时候终止在某表中的查找
    FirstFindRangeCol: integer; //在某一工作表中找到的第一个单元格的列号，用于在适当时候终止在某表中的查找

    FFindSearchLookin: TSearchLookin;
    FCanShowCombo: Boolean;

    FSwitchFindReplaceMode: array[TFindReplaceMode] of procedure of object;
    FSwitchDialogMode: array[TDialogMode] of procedure of object;
    FApplication: IETApplication;

    FbProtected: Boolean;

    function GetFindLookin: TOleEnum;
    function GetFindLookAt: TOleEnum;
    function GetSearchOrder: TOleEnum;
    function GetMatchCase: OleVariant;
    function GetMatchByte: OleVariant;

    function Find(FRange: IETRange): IETRange;
    function Replace(FRange: IETRange; var fIsReplaceFail: Boolean): IETRange;
    function ReplaceAll(FRange: IETRange): Integer;

    function IsSingleCell(FRange: IETRange): Boolean; //单个单元格、合并单元格都算是单个单元格

    procedure FormActivate(Sender: TObject);
    procedure EnableDataControl(const bEnabled: Boolean);
    function GetCellType: ETCellType;
    function GetDataType: ETSpecialCellsValue;
    procedure SetCellType(const Value: ETCellType);
    procedure SetDataType(const Value: ETSpecialCellsValue);
    procedure SwitchMode(const tabControl: TTntTabControl);
    procedure FindAll;
    procedure DeleteRepeat(I: Integer);
    function colSetWidth(I: Integer): TWidth;
  protected
    procedure CreateWnd; override;
    procedure CreateParams(var Params: TCreateParams); override;
  protected
    class function SD_GetType_: TDialogType; override;
    procedure SD_ObjParamChange; override;
    procedure SD_EnvParamChange; override;
    function SD_Update: HResult; override;

    procedure DoShow; override;
    procedure DoClose(var Action: TCloseAction); override;
  public
    { Public declarations }
    function IsShortCut(var Message: TWMKey): Boolean; override;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure SaveState;
    procedure RestoreState;

    procedure SwitchToFind;
    procedure SwitchToReplace;
    procedure SwitchToNormal;
    procedure SwitchToAdvanced;
    procedure SwitchToGoto;
  end;
var
  Form1: TFindReplaceDlg;
  var ColumnToSort: Integer;
  var asc: Boolean;
  var WorkSheetStr: WideString;

implementation

uses
  ETShellID, ETConsts, MainDm, kso10
  {TNT-USE}
  , WideStrUtils, TntMenus, Math
  {TNT-END};

{$IFNDEF ExeExcludeRes}
{$R *.dfm}
{$ENDIF}


{$I etshell.inc}

const
  iHeightNormalMode = 156;
  iHeightAdvancedMode = 204;
  iWidthAnyMode = 455;

const
  etCellFormulasAndConstants = 25;
  etCellCurDataArea = 26;
  etCellHyperlink = 27;

var
  FindReplaceMode: TFindReplaceMode;
  DialogMode: TDialogMode = dmNormal;
  LastDialogMode: TDialogMode = dmAdvanced;
  SearchRange: TSearchRange = srSheet;
  SearchOrder: TSearchOrder = soByRows;
  SearchLookin: TSearchLookin = slFormulas;
  SearchOptions: TSearchOptions = [];
  FindContent: WideString;
  ReplaceContent: WideString;
  CellType: ETCellType;
  DataValue: ETCellType;
  DataType: OleVariant;
  GoToObject: Boolean = False;

(* --------------------------------------------------------------------------------- *)
(*
(* Class:
(*   Create by Yeppy.Arch.WPS Date: 2004-3-25
(*
(*   Category:
(*   Inherited:
(*   Comment:
(* --------------------------------------------------------------------------------- *)

{ ---------------------------------------------------------------------------------- }
{ Comment by Yeppy.Arch.WPS Date: 2004-3-25
{   public
{ ---------------------------------------------------------------------------------- }

procedure TFindReplaceDlg.FormDeactivate(Sender: TObject);
begin
  //当发现对话框失去焦点，则下一个视为重新开始新的查找替换
  _FindNextFlag := false;
  _ReplaceFirstFlag := true;
end;

procedure TFindReplaceDlg.FormResize(Sender: TObject);
begin
//  with Constraints do
//    gpResizer.DrawHandler := (Width <> MinWidth) or (Height <> MinHeight);
  lvFindAll.Height := MAX(ClientHeight - lvFindAll.Top - ScaleUISizeY(8), 0);

end;

procedure TFindReplaceDlg.SwitchMode(const tabControl: TTntTabControl);
begin
  FSwitchFindReplaceMode[FindReplaceMode];
  if FindReplaceMode = frmGoto then
  begin
    FSwitchDialogMode[dmAdvanced];
  end else
    FSwitchDialogMode[DialogMode];
  Caption := WideStripHotkey(tabControl.Tabs[tabControl.TabIndex]);
end;

procedure TFindReplaceDlg.tcFindReplaceChange(Sender: TObject);
var
  tabControl: TTntTabControl;
begin
  tabControl := Sender as TTntTabControl;
  FindReplaceMode := TFindReplaceMode(tabControl.TabIndex);
  SwitchMode(tabControl);
end;

procedure TFindReplaceDlg.tcFindReplaceEnter(Sender: TObject);
begin
  // FSwitchFindReplaceMode[FindReplaceMode];
end;

procedure TFindReplaceDlg.mcbFindChange(Sender: TObject);
begin
  _FindNextFlag := False;
end;

procedure TFindReplaceDlg.mcbReplaceChange(Sender: TObject);
begin
  _FindNextFlag := False;
end;

procedure TFindReplaceDlg.cbRangeSelect(Sender: TObject);
begin
  _FindNextFlag := False;
  SearchRange := TSearchRange(cbRange.ItemIndex);
end;

procedure TFindReplaceDlg.cbOrderSelect(Sender: TObject);
begin
  _FindNextFlag := False;
  SearchOrder := TSearchOrder(cbOrder.ItemIndex);
end;

procedure TFindReplaceDlg.cbLookinSelect(Sender: TObject);
begin
  _FindNextFlag := False;
  SearchLookin := TSearchLookin(cbLookin.ItemIndex);
end;

procedure TFindReplaceDlg.cbMatchCaseClick(Sender: TObject);
begin
  _FindNextFlag := False;
  if cbMatchCase.Checked then
    SearchOptions := SearchOptions + [soMatchCase]
  else
    SearchOptions := SearchOptions - [soMatchCase];
end;

procedure TFindReplaceDlg.cbMatchContentsClick(Sender: TObject);
begin
  _FindNextFlag := False;
  if cbMatchContents.Checked then
    SearchOptions := SearchOptions + [soMatchContents]
  else
    SearchOptions := SearchOptions - [soMatchContents];
end;

procedure TFindReplaceDlg.cbMatchByteClick(Sender: TObject);
begin
  _FindNextFlag := False;
  if cbMatchByte.Checked then
    SearchOptions := SearchOptions + [soMatchByte]
  else
    SearchOptions := SearchOptions - [soMatchByte];
end;

procedure TFindReplaceDlg.btnCloseClick(Sender: TObject);
begin
  ModalResult := mrCancel;
  Close;
end;

procedure TFindReplaceDlg.btnOptionsClick(Sender: TObject);
begin
  DialogMode := TDialogMode((Integer(DialogMode) + 1) mod 2);
  FSwitchDialogMode[DialogMode];
end;

procedure TFindReplaceDlg.btnFindNextClick(Sender: TObject);
var
  Selection: IDispatch;
  ActiveSheet: IDispatch;
  NextSheet: IDispatch; //遍历中使用
  WorkSheet: IWorkSheet;
  FindRes: IETRange;
  NoFindOutSheetCount: Integer; //查找过但没有找到待查内容的工作表数目，用于工作簿和多选工作表时的查找
  SheetIndex: Integer;
  SelSheets: Sheets;
begin
  // todo: 这句用于去除警告，应该仔细检查逻辑关系。
  SheetIndex := 0;

  if FApplication = nil then
    Exit;
  if Assigned(FApplication) then
    SelSheets := FApplication.ActiveWindow.SelectedSheets;

  Selection := FApplication.Selection[$0804];
  if (Selection = nil) then exit;

  //如果待查找内容为空，提示错误
  if mcbFind.Text = '' then begin
    KSO_MessageBox(WideLoadResString(@sFindReplace_InvalidWhat), '', MB_OK + MB_ICONEXCLAMATION);
    mcbFind.SetFocus;
    Exit;
  end;

  { add by oyxz, see #C1 }
  // 从上面移过来的代码，如果为空值时，不进行以下操作，否则会导致bug:26096
  mcbFind.ActivateCurrItem;
  mcbReplace.ActivateCurrItem;
  { add end }

  //如果当前所选不是区域，则什么也不做
  if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
    Exit;
  end;

  //确定查找范围，注意有的情况不必改变
  //对于工作表，第一次查找之后都不必改变
  if SearchRange = srWorkBook then begin
    //如果当前查找范围是个工作簿

    //如果是第一次查找，则获取当前工作表，把Cells付给查找范围，而且要清除当前的Selection
    if not _FindNextFlag then begin
      SheetsCount := FApplication.ActiveWorkbook.Worksheets.Count;
      _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
    end;

    ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
    if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
      FindReplaceRange := WorkSheet.Cells;
       //FApplication.ActiveCell.Select;//取消原工作表上的区域选择
    end
    else begin
      Exit;
    end;

    //（隐含一点，默认拿当前查找范围进行查找，是上次留下的）
    FindRes := Find(FindReplaceRange);

    //处理是否换表查找
    //底下也要这么做  （有机会做成一个函数）
    if FindRes <> nil then begin
      //如果找到了
      if _FirstFindInSheet then begin
        //若是在当前工作表中第一次发现
        //把标志设为假，并把找到的内容保存起来
        _FirstFindInSheet := false;
        FirstFindRangeRow := FindRes.Row;
        FirstFindRangeCol := FindRes.Column;
      end
      else begin
        //如果不是，则比较找到的和保存的
        if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
           //如果一致，说明找了一轮了，要进行下一次查找，把查找结果赋空
          FindRes := nil;
        end;
      end;
    end;

    if FindRes = nil then begin
       //没找到待查内容的工作表数目初始化为0
      NoFindOutSheetCount := 0;
       //获取当前工作表索引为工作表索引
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //没找到待查内容工作表数目已经为整个工作表数目
             //说明整个工作簿都没有想要的内容，退出查找
          break;
        end
        else begin
             //找到下一个工作表（将工作表索引加1取模）
             //模为工作表总数加1，因为工作表索引从1开始。若取模结果为0，要调整为1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := FApplication.ActiveWorkbook.Worksheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
            FindReplaceRange := WorkSheet.Cells;
               //换了工作表，要调整选中状态，否则当前活动单元格还在上一个表上
               //会影响查找的起始位置
            _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
          end;
        end;

        if WorkSheet.Visible <> etSheetVisible then
          FindRes := nil // 如果工作表不可见，则不查找
        else
          FindRes := Find(FindReplaceRange);

          //判断是否对一个表查找了一轮
        if FindRes <> nil then begin
            //如果找到了
            //只有当该工作表有匹配内容时，才将该表选中
          WorkSheet.Select(true);
            //查找两遍的原因：
            //1、如果一开始就选中工作表，则该工作表若没有被查找内容就会出现闪烁（该表一闪而过）
            //2、如果开始不选中工作表查找，则当前单元格将会是上个工作表的，导致新表查找位置有问题
            //所以采用的方法是先找，如果找到，则选中再找
          Find(FindReplaceRange);
          if _FirstFindInSheet then begin
              //若是在当前工作表中第一次发现
              //把标志设为假，并把找到的内容保存起来
            _FirstFindInSheet := false;
            FirstFindRangeRow := FindRes.Row;
            FirstFindRangeCol := FindRes.Column;
          end
          else begin
              //如果不是，则比较找到的和保存的
            if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
                 //如果一致，说明找了一轮了，要进行下一次查找，把查找结果赋空
              FindRes := nil;
            end;
          end;
        end;

        if FindRes = nil then
             //当发现新选取的一个工作表没有找到，则将遍历表数目加1，表示该表没有待查找内容
             //循环外的查找失败不能证明整表无查找内容，有可能是查到尾了
          NoFindOutSheetCount := NoFindOutSheetCount + 1;
        begin
        end
      until (FindRes <> nil);

      if FindRes <> nil then begin
        FindRes.Select;
      end;
    end;
  end
  ////////////////////////////////////////////////////////////////////////////////////
  else if SelSheets.Count > 1 then begin
    //如果当前选中了多个工作表
    if not _FindNextFlag then begin
      _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
      //此处记录的工作表数目是当前选中的工作表数目
      SheetsCount := SelSheets.Count;
    end;

    //每次查找都要重新确定一下查找范围
    //和单表查找一样，如果当前表内所选单元格数目为1，则进行整表查找
    //否则只是对Selection进行查找
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end;
    end;

    //（隐含一点，默认拿当前查找范围进行查找，是上次留下的）
    FindRes := Find(FindReplaceRange);

    //处理是否换表查找
    //底下也要这么做  （有机会做成一个函数）
    if FindRes <> nil then begin
      //如果找到了
      if _FirstFindInSheet then begin
        //若是在当前工作表中第一次发现
        //把标志设为假，并把找到的内容保存起来
        _FirstFindInSheet := false;
        FirstFindRangeRow := FindRes.Row;
        FirstFindRangeCol := FindRes.Column;
      end
      else begin
        //如果不是，则比较找到的和保存的
        if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
           //如果一致，说明找了一轮了，要进行下一次查找，把查找结果赋空
          FindRes := nil;
        end;
      end;
    end;

    if FindRes = nil then begin
       //没找到待查内容的工作表数目初始化为0
      NoFindOutSheetCount := 0;
       //获取当前工作表索引为工作表索引
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //没找到待查内容工作表数目已经为整个工作表数目
             //说明整个工作簿都没有想要的内容，退出查找
          break;
        end
        else begin
             //找到下一个工作表（将工作表索引加1取模）
             //模为工作表总数加1，因为工作表索引从1开始。若取模结果为0，要调整为1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := SelSheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
                //选了一个新表，重新调整查找范围
                //先选择新表，然后调整范围
                //换了工作表，要调整选中状态，否则当前活动单元格还在上一个表上
            WorkSheet.Activate;
            WorkSheet.Select(false);

                //如果当前所选不是区域，则什么也不做
            Selection := FApplication.Selection[$0804];
            if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
              Exit;
            end;

                //和单表查找一样，如果当前表内所选单元格数目为1，则进行整表查找
                //否则只是对Selection进行查找
            if FindReplaceRange.Count = 1 then begin
              FindReplaceRange := WorkSheet.Cells;
            end;

                //会影响查找的起始位置
            _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
          end;
        end;

        FindRes := Find(FindReplaceRange);

          //判断是否对一个表查找了一轮
        if FindRes <> nil then begin
            //如果找到了
          if _FirstFindInSheet then begin
              //若是在当前工作表中第一次发现
              //把标志设为假，并把找到的内容保存起来
            _FirstFindInSheet := false;
            FirstFindRangeRow := FindRes.Row;
            FirstFindRangeCol := FindRes.Column;
          end
          else begin
              //如果不是，则比较找到的和保存的
            if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
                 //如果一致，说明找了一轮了，要进行下一次查找，把查找结果赋空
              FindRes := nil;
            end;
          end;
        end;

        if FindRes = nil then
             //当发现新选取的一个工作表没有找到，则将遍历表数目加1，表示该表没有待查找内容
             //循环外的查找失败不能证明整表无查找内容，有可能是查到尾了
          NoFindOutSheetCount := NoFindOutSheetCount + 1;
        begin
        end
      until (FindRes <> nil);

//       if FindRes <> nil then
//       begin
//          //对于多选工作表，遍历到一个新的工作表，并不取消以前工作表的选中
//          WorkSheet.Activate;
//       end;
    end;
  end
  else begin
    //如果当前只选中了一个工作表，在第一次查找的时候需要确定查找范围
    //如果选中的是一个单元格，需要把整表作为待查范围
    //如果选中的是两个以上单元格，当前的选择区域就是待查范围，不必额外处理
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end
    end;

    //实际进行查找
    FindRes := Find(FindReplaceRange);
  end;

  //对查找到的结果做处理
  if FindRes <> nil then begin
    FindRes.Activate;
  end
  else begin
    KSO_MessageBox(WideLoadResString(@sFindReplace_NotFind), '', MB_ICONEXCLAMATION + MB_OK);
    mcbFind.SetFocus;
    _FindCompleteFlag := true;
  end;
end;

procedure TFindReplaceDlg.btnFindAllClick(Sender: TObject);
begin
  lvFindAll.Items.BeginUpdate;
  try
    FindAll;
  finally
    lvFindAll.Items.EndUpdate;
    lvFindAll.Columns[lvFindAll.Columns.Count - 1].Width := lvFindAll.ClientWidth -
            colSetWidth(lvFindAll.Columns.Count - 1);
  end;
  if (lvFindAll.Selected = nil) and (lvFindAll.Items.Count > 0) then
    lvFindAll.Items[0].Selected := True;
  lvFindAll.SetFocus;
end;
function TFindReplaceDlg.colSetWidth(I: Integer):TWidth;
begin
  if I = 1 then begin
    Result := lvFindAll.Columns[0].Width;
  end else
    Result := colSetWidth(I - 1) + lvFindAll.Columns[I - 1].Width;
end;
procedure TFindReplaceDlg.DeleteRepeat(I: Integer);
var
  J: Integer;
begin
  if I = 0 then
    Exit;
  for  J := 0 to I - 1 do
  begin
    if lvFindAll.Items[J].Caption = lvFindAll.Items[I].Caption then
      if lvFindAll.Items[J].SubItems.Strings[0] = lvFindAll.Items[I].SubItems.Strings[0] then
      if lvFindAll.Items[J].SubItems.Strings[2] = lvFindAll.Items[I].SubItems.Strings[2] then begin
        lvFindAll.Items[I].Delete;
        DeleteRepeat(lvFindAll.Items.Count - 1);
        Break;
      end;

  end;

end;  
procedure TFindReplaceDlg.FindAll;
var
  Selection: IDispatch;
  ActiveSheet: IDispatch;
  NextSheet: IDispatch; //遍历中使用
  WorkSheet: IWorkSheet;
  FindRes: IETRange;
  RangeTemp: OleVariant;
  NoFindOutSheetCount: Integer; //查找过但没有找到待查内容的工作表数目，用于工作簿和多选工作表时的查找
  SheetIndex: Integer;
  SelSheets: Sheets;
  listitem: TTntListItem;
  NameTemp: IName;
  AddressStr :Widestring;
  AddressTemp :WideString;
  I: Integer; //循环变量
begin
  // todo: 这句用于去除警告，应该仔细检查逻辑关系。
  SheetIndex := 0;
  if FApplication = nil then
    Exit;
  if Assigned(FApplication) then
    SelSheets := FApplication.ActiveWindow.SelectedSheets;

  Selection := FApplication.Selection[$0804];
  if (Selection = nil) then exit;

  //如果待查找内容为空，提示错误
  if mcbFind.Text = '' then begin
    lvFindAll.Clear;
    KSO_MessageBox(WideLoadResString(@sFindReplace_InvalidWhat), '', MB_OK + MB_ICONEXCLAMATION);
    mcbFind.SetFocus;
    Exit;
  end;

  { add by oyxz, see #C1 }
  // 从上面移过来的代码，如果为空值时，不进行以下操作，否则会导致bug:26096
  mcbFind.ActivateCurrItem;
  mcbReplace.ActivateCurrItem;
  { add end }

  //如果当前所选不是区域，则什么也不做
  if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
    Exit;
  end;

  //确定查找范围，注意有的情况不必改变
  //对于工作表，第一次查找之后都不必改变
  if SearchRange = srWorkBook then begin
    //如果当前查找范围是个工作簿
    //如果是第一次查找，则获取当前工作表，把Cells付给查找范围，而且要清除当前的Selection
    if not _FindNextFlag then begin
      SheetsCount := FApplication.ActiveWorkbook.Worksheets.Count;
      _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
    end;

    ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
    if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
      FindReplaceRange := WorkSheet.Cells;
       //FApplication.ActiveCell.Select;//取消原工作表上的区域选择
    end
    else begin
      Exit;
    end;

    //（隐含一点，默认拿当前查找范围进行查找，是上次留下的）
    FindRes := Find(FindReplaceRange);

    //处理是否换表查找
    //底下也要这么做  （有机会做成一个函数）
    if FindRes <> nil then begin
      //如果找到了
      if _FirstFindInSheet then begin
        //若是在当前工作表中第一次发现
        //把标志设为假，并把找到的内容保存起来
        _FirstFindInSheet := false;
        FirstFindRangeRow := FindRes.Row;
        FirstFindRangeCol := FindRes.Column;
      end
      else begin
        //如果不是，则比较找到的和保存的
        if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
           //如果一致，说明找了一轮了，要进行下一次查找，把查找结果赋空
          FindRes := nil;
        end;
      end;
    end;

    if FindRes = nil then begin
       //没找到待查内容的工作表数目初始化为0
      NoFindOutSheetCount := 0;
       //获取当前工作表索引为工作表索引
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //没找到待查内容工作表数目已经为整个工作表数目
             //说明整个工作簿都没有想要的内容，退出查找
          break;
        end
        else begin
             //找到下一个工作表（将工作表索引加1取模）
             //模为工作表总数加1，因为工作表索引从1开始。若取模结果为0，要调整为1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := FApplication.ActiveWorkbook.Worksheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
            FindReplaceRange := WorkSheet.Cells;
               //换了工作表，要调整选中状态，否则当前活动单元格还在上一个表上
               //会影响查找的起始位置
            _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
          end;
        end;

        if WorkSheet.Visible <> etSheetVisible then
          FindRes := nil // 如果工作表不可见，则不查找
        else
          FindRes := Find(FindReplaceRange);

          //判断是否对一个表查找了一轮
        if FindRes <> nil then begin
            //如果找到了
            //只有当该工作表有匹配内容时，才将该表选中
          WorkSheet.Select(true);
            //查找两遍的原因：
            //1、如果一开始就选中工作表，则该工作表若没有被查找内容就会出现闪烁（该表一闪而过）
            //2、如果开始不选中工作表查找，则当前单元格将会是上个工作表的，导致新表查找位置有问题
            //所以采用的方法是先找，如果找到，则选中再找
          Find(FindReplaceRange);
          if _FirstFindInSheet then begin
              //若是在当前工作表中第一次发现
              //把标志设为假，并把找到的内容保存起来
            _FirstFindInSheet := false;
            FirstFindRangeRow := FindRes.Row;
            FirstFindRangeCol := FindRes.Column;
          end
          else begin
              //如果不是，则比较找到的和保存的
            if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
                 //如果一致，说明找了一轮了，要进行下一次查找，把查找结果赋空
              FindRes := nil;
            end;
          end;
        end;

        if FindRes = nil then
             //当发现新选取的一个工作表没有找到，则将遍历表数目加1，表示该表没有待查找内容
             //循环外的查找失败不能证明整表无查找内容，有可能是查到尾了
          NoFindOutSheetCount := NoFindOutSheetCount + 1;
        begin
        end
      until (FindRes <> nil);

      if FindRes <> nil then begin
        FindRes.Select;
      end;
    end;
  end
  ////////////////////////////////////////////////////////////////////////////////////
  else if SelSheets.Count > 1 then begin
    //如果当前选中了多个工作表
    if not _FindNextFlag then begin
      _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
      //此处记录的工作表数目是当前选中的工作表数目
      SheetsCount := SelSheets.Count;
    end;

    //每次查找都要重新确定一下查找范围
    //和单表查找一样，如果当前表内所选单元格数目为1，则进行整表查找
    //否则只是对Selection进行查找
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end;
    end;

    //（隐含一点，默认拿当前查找范围进行查找，是上次留下的）
    FindRes := Find(FindReplaceRange);

    //处理是否换表查找
    //底下也要这么做  （有机会做成一个函数）
    if FindRes <> nil then begin
      //如果找到了
      if _FirstFindInSheet then begin
        //若是在当前工作表中第一次发现
        //把标志设为假，并把找到的内容保存起来
        _FirstFindInSheet := false;
        FirstFindRangeRow := FindRes.Row;
        FirstFindRangeCol := FindRes.Column;
      end
      else begin
        //如果不是，则比较找到的和保存的
        if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
           //如果一致，说明找了一轮了，要进行下一次查找，把查找结果赋空
          FindRes := nil;
        end;
      end;
    end;

    if FindRes = nil then begin
       //没找到待查内容的工作表数目初始化为0
      NoFindOutSheetCount := 0;
       //获取当前工作表索引为工作表索引
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //没找到待查内容工作表数目已经为整个工作表数目
             //说明整个工作簿都没有想要的内容，退出查找
          break;
        end
        else begin
             //找到下一个工作表（将工作表索引加1取模）
             //模为工作表总数加1，因为工作表索引从1开始。若取模结果为0，要调整为1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := SelSheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
                //选了一个新表，重新调整查找范围
                //先选择新表，然后调整范围
                //换了工作表，要调整选中状态，否则当前活动单元格还在上一个表上
            WorkSheet.Activate;
            WorkSheet.Select(false);

                //如果当前所选不是区域，则什么也不做
            Selection := FApplication.Selection[$0804];
            if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
              Exit;
            end;

                //和单表查找一样，如果当前表内所选单元格数目为1，则进行整表查找
                //否则只是对Selection进行查找
            if FindReplaceRange.Count = 1 then begin
              FindReplaceRange := WorkSheet.Cells;
            end;

                //会影响查找的起始位置
            _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
          end;
        end;

        FindRes := Find(FindReplaceRange);

          //判断是否对一个表查找了一轮
        if FindRes <> nil then begin
            //如果找到了
          if _FirstFindInSheet then begin
              //若是在当前工作表中第一次发现
              //把标志设为假，并把找到的内容保存起来
            _FirstFindInSheet := false;
            FirstFindRangeRow := FindRes.Row;
            FirstFindRangeCol := FindRes.Column;
          end
          else begin
              //如果不是，则比较找到的和保存的
            if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
                 //如果一致，说明找了一轮了，要进行下一次查找，把查找结果赋空
              FindRes := nil;
            end;
          end;
        end;

        if FindRes = nil then
             //当发现新选取的一个工作表没有找到，则将遍历表数目加1，表示该表没有待查找内容
             //循环外的查找失败不能证明整表无查找内容，有可能是查到尾了
          NoFindOutSheetCount := NoFindOutSheetCount + 1;
        begin
        end
      until (FindRes <> nil);

//       if FindRes <> nil then
//       begin
//          //对于多选工作表，遍历到一个新的工作表，并不取消以前工作表的选中
//          WorkSheet.Activate;
//       end;
    end;
  end
  else begin
    //如果当前只选中了一个工作表，在第一次查找的时候需要确定查找范围
    //如果选中的是一个单元格，需要把整表作为待查范围
    //如果选中的是两个以上单元格，当前的选择区域就是待查范围，不必额外处理
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end
    end;

    //实际进行查找
    FindRes := Find(FindReplaceRange);
  end;

  //对查找到的结果做处理
  if FindRes <> nil then begin
    AddressStr := Findres.Address[True, True ,etA1, True, RangeTemp];
    if self.FindFirst = nil then begin
      lvFindAll.Items.Clear;
      self.FindFirst := FindRes;
    end else if FindFirst.Address[True, True ,etA1, True, RangeTemp] = AddressStr then begin
      self.FindFirst := nil;
      Exit;
    end;

    listitem := lvFindAll.Items.Add;
    listitem.Caption := FindRes.Worksheet.Application.ActiveWorkbook.Name;
    listitem.SubItems.Add(FindRes.Worksheet.Name);

    for I := FindRes.Worksheet.Application.Names.Count downto 1 do
    begin
      NameTemp := FindRes.Worksheet.Application.Names.Item(I,RangeTemp,RangeTemp);
      AddressTemp := '=' + FindRes.Worksheet.Name + '!' + Copy(AddressStr,Length(AddressStr)-3,4);
      if NameTemp.RefersTo = AddressTemp then begin
        AddressTemp := NameTemp.Name_;
        Break;
      end
      else begin
        AddressTemp := '';
      end;
    end;

    listitem.SubItems.Add(AddressTemp);
    listitem.SubItems.Add(Copy(AddressStr,Length(AddressStr)-3,4));
    listitem.SubItems.Add(FindRes.Text);
    if FindRes.Text <> FindRes.Formula then
    begin
      listitem.SubItems.Add(FindRes.Formula);
    end
    else begin
      listitem.SubItems.Add('');
    end;

    FindAll;
  end
  else begin
    lvFindAll.Clear;//清空表
    KSO_MessageBox(WideLoadResString(@sFindReplace_NotFind), '', MB_ICONEXCLAMATION + MB_OK);
    mcbFind.SetFocus;
    _FindCompleteFlag := true;
  end;

  if lvFindAll.Height < ScaleUISizeY(30) then
    ClientHeight := ClientHeight + ScaleUISizeY(150) - lvFindAll.Height;

  if lvFindAll.Items.Count <> 0 then begin
    if SearchRange = srWorkBook then
      DeleteRepeat(lvFindAll.Items.Count - 1);
    asc := False;
    ColumnToSort := 3;
    lvFindAll.AlphaSort;
    ColumnToSort := 1;
    lvFindAll.AlphaSort;
    ActiveSheet := FApplication.ActiveWorkbook.Worksheets[lvFindAll.Items[0].SubItems.Strings[0]];
    if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
      WorkSheet.Activate;
      WorkSheet.Range[lvFindAll.Items[0].SubItems.Strings[2],lvFindAll.Items[0].SubItems.Strings[2]].Select;
    end;
  end;
end;
    
procedure TFindReplaceDlg.btnReplaceClick(Sender: TObject);
var
  Selection: IDispatch;
  ActiveSheet: IDispatch;
  WorkSheet: IWorkSheet;
  FindRes: IETRange;
  NoFindOutSheetCount: Integer; //查找过但没有找到待查内容的工作表数目，用于工作簿和多选工作表时的查找
  SheetIndex: Integer;
  NextSheet: IDispatch; //遍历中使用
  fIsReplaceFail: Boolean;
  SelSheets: Sheets;
begin
  SheetIndex := 1; //这样做意义并不大，只是为了少出现一个“might not have been initialized”的Warning
  fIsReplaceFail := false;

  Selection := FApplication.Selection[$0804];
  if (Selection = nil) then exit;

  if Assigned(FApplication) then
    SelSheets := FApplication.ActiveWindow.SelectedSheets;
  if _ReplaceFirstFlag then begin
    if FApplication = nil then
      Exit;

    mcbFind.ActivateCurrItem;
    mcbReplace.ActivateCurrItem;

    //如果待查找内容为空，提示错误
    if mcbFind.Text = '' then begin
      KSO_MessageBox(WideLoadResString(@sFindReplace_InvalidWhat), '', MB_OK + MB_ICONEXCLAMATION);
      mcbFind.SetFocus;
      Exit;
    end;
  end;

  //如果当前所选不是区域，则什么也不做
  if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
    Exit;
  end;

  //确定查找范围，注意有的情况不必改变
  //对于工作表，第一次查找之后都不必改变
  //
  if SearchRange = srWorkBook then begin
    //如果当前查找范围是个工作簿
    //如果是第一次查找，则获取当前工作表，把Cells付给查找范围，而且要清除当前的Selection
    if _ReplaceFirstFlag then begin
      SheetsCount := FApplication.ActiveWorkbook.Worksheets.Count;
      _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
    end;

    ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
    if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
      FindReplaceRange := WorkSheet.Cells;
      FApplication.ActiveCell.Select; //取消原工作表上的区域选择
    end
    else begin
      Exit;
    end;

    if WorkSheet.Visible <> etSheetVisible then
      FindRes := nil  // 如果工作表被隐藏不替换
    else
      //（隐含一点，默认拿当前查找范围进行查找，是上次留下的）
      FindRes := Replace(FindReplaceRange, fIsReplaceFail);

    if FindRes = nil then begin
       //如果在当前工作表没有找到，则要考虑到下一个工作表里查找
       //没找到待查内容的工作表数目初始化为0
      NoFindOutSheetCount := 0;
       //获取当前工作表索引
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //在工作簿中循环查找
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //没找到待查内容工作表数目已经为整个工作表数目
             //说明整个工作簿都没有想要的内容，退出查找
          break;
        end
        else begin
             //找到下一个工作表（将工作表索引加1取模）
             //模为工作表总数加1，因为工作表索引从1开始。若取模结果为0，要调整为1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := FApplication.ActiveWorkbook.Worksheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
            FindReplaceRange := WorkSheet.Cells;
               //换了工作表，要调整选中状态，否则当前活动单元格还在上一个表上
               //会影响查找的起始位置
            _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
          end;
        end;

        if WorkSheet.Visible <> etSheetVisible then
          FindRes := nil  // 如果工作表被隐藏不替换
        else
          FindRes := Find(FindReplaceRange);

        if FindRes = nil then begin
             //当发现新选取的一个工作表没有找到，则将遍历表数目加1，表示该表没有待查找内容
             //循环前的查找失败不能证明整表无查找内容，有可能第一个工作表查到尾了
          NoFindOutSheetCount := NoFindOutSheetCount + 1;
        end
        else begin
             //只有当该工作表有匹配内容时，才将该表选中
          WorkSheet.Select(true);
             //再找一遍，原因和工作簿查找同理
          FindRes := Find(FindReplaceRange);
        end
      until (FindRes <> nil);
    end;
  end
////////////////////////////////////////////////////////////////////////////
  else if SelSheets.Count > 1 then begin
    //如果当前选中了多个工作表
    if _ReplaceFirstFlag then begin
      SheetsCount := SelSheets.Count;
      _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
    end;

    //每次查找都要重新确定一下查找范围
    //和单表查找一样，如果当前表内所选单元格数目为1，则进行整表查找
    //否则只是对Selection进行查找
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end;
    end;

    //（隐含一点，默认拿当前查找范围进行查找，是上次留下的）
    FindRes := Replace(FindReplaceRange, fIsReplaceFail);

    if FindRes = nil then begin
       //如果在当前工作表没有找到，则要考虑到下一个工作表里查找
       //没找到待查内容的工作表数目初始化为0
      NoFindOutSheetCount := 0;
       //获取当前工作表索引
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet是个Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //在工作簿中循环查找
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //没找到待查内容工作表数目已经为整个工作表数目
             //说明整个工作簿都没有想要的内容，退出查找
          break;
        end
        else begin
             //找到下一个工作表（将工作表索引加1取模）
             //模为工作表总数加1，因为工作表索引从1开始。若取模结果为0，要调整为1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := SelSheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
                //选了一个新表，重新调整查找范围
                //先选择新表，然后调整范围
                //换了工作表，要调整选中状态，否则当前活动单元格还在上一个表上
            WorkSheet.Activate;
            WorkSheet.Select(false);

                //如果当前所选不是区域，则什么也不做
            Selection := FApplication.Selection[$0804];
            if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
              Exit;
            end;

                //和单表查找一样，如果当前表内所选单元格数目为1，则进行整表查找
                //否则只是对Selection进行查找
            if FindReplaceRange.Count = 1 then begin
              FindReplaceRange := WorkSheet.Cells;
            end;

                //会影响查找的起始位置
            _FirstFindInSheet := true; //每次改换sheet后都要把该标志设为TRUE
          end;
        end;

        FindRes := Find(FindReplaceRange);

        if FindRes = nil then
             //当发现新选取的一个工作表没有找到，则将遍历表数目加1，表示该表没有待查找内容
             //循环前的查找失败不能证明整表无查找内容，有可能第一个工作表查到尾了
          NoFindOutSheetCount := NoFindOutSheetCount + 1;
        begin
        end
      until (FindRes <> nil);
    end;
  end
  else begin
    //如果当前只选中了一个工作表
    //如果选中的是一个单元格，需要把整表作为待查范围
    //如果选中的是两个以上单元格，当前的选择区域就是待查范围，不必额外处理
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet;
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end
    end;

    //实际进行替换
    FindRes := Replace(FindReplaceRange, fIsReplaceFail);
  end;

  if FindRes <> nil then
    FindRes.Activate
  else begin
    if FbProtected then
      // 保护时不弹框，内核已经弹过了
    else if fIsReplaceFail = true then
      //替换过程中发现问题
      KSO_MessageBox(WideLoadResString(@sReplace_ReplaceFailed), '', MB_OK + MB_ICONEXCLAMATION)
    else
      //找不到匹配项
      KSO_MessageBox(WideLoadResString(@sReplace_NotFind), '', MB_OK + MB_ICONEXCLAMATION);
    mcbFind.SetFocus;
    end;
end;

procedure TFindReplaceDlg.btnReplaceAllClick(Sender: TObject);
var
  Selection: IDispatch;
  ActiveSheet: IDispatch;
  NextSheet: IDispatch;
  WorkSheet: IWorkSheet;
  ReplaceAllCount: Integer; //总的全部替换的数目
  SheetReplaceAllCount: Integer; //工作表全部替换的数目
  SheetsCount: Integer; //工作簿内工作表的数目
  I: Integer; //循环变量
  AllSheets: Worksheets;  // cache it
  SelSheets: Sheets;
begin
  ReplaceAllCount := 0;

  if FApplication = nil then
    Exit;
  AllSheets := FApplication.ActiveWorkbook.Worksheets;
  SelSheets := FApplication.ActiveWindow.SelectedSheets;
  Selection := FApplication.Selection[$0804];
  if (Selection = nil) then exit;

  mcbFind.ActivateCurrItem;
  mcbReplace.ActivateCurrItem;

  //如果待查找内容为空，提示错误
  if mcbFind.Text = '' then begin
    KSO_MessageBox(WideLoadResString(@sFindReplace_InvalidWhat), '', MB_OK + MB_ICONEXCLAMATION);
    mcbFind.SetFocus;
    Exit;
  end;

  //如果当前所选不是区域，则什么也不做
  if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
    Exit;
  end;

  //确定查找范围
  if SearchRange = srWorkBook then begin
    //如果当前查找范围是个工作簿
    //遍历各表，逐表全部替换，记录替换的总数

    //确定工作簿内有几个工作表，初始化查找替换结果
    SheetsCount := AllSheets.Count;
    for I := 1 to SheetsCount do begin
      //获取工作表
      NextSheet := AllSheets[I];
      if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        if WorkSheet.Visible <> etSheetVisible then continue;
		if WorkSheet.ProtectContents then
               continue;
          //将整个工作表作为查找范围
        FindReplaceRange := WorkSheet.Cells;
          //对该工作表进行全部替换
        SheetReplaceAllCount := ReplaceAll(FindReplaceRange);
        if SheetReplaceAllCount < 0 then begin
          ReplaceAllCount := SheetReplaceAllCount;
          break;
        end
        else begin
          ReplaceAllCount := ReplaceAllCount + SheetReplaceAllCount;
        end;
      end
      else
        Exit;
    end;
  end
  else if SelSheets.Count > 1 then begin
    //如果当前选中了多个工作表
    SheetsCount := SelSheets.Count;
    for I := 1 to SheetsCount do begin
      //获取工作表
      NextSheet := SelSheets[I];
      if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
		if WorkSheet.ProtectContents then
                continue;
        //换表
        WorkSheet.Activate;
        WorkSheet.Select(false);

        //重取selection（奇怪，如果不这样做，第一次可以拿到，第二次这个区域的数目就为0了）
        Selection := FApplication.Selection[$0804];
        if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
          Exit;
        end;

        if IsSingleCell(FindReplaceRange) then begin
           //将整个工作表作为查找范围
          FindReplaceRange := WorkSheet.Cells;
        end
        else begin
           //如果当前所选是区域
          Selection := FApplication.Selection[$0804];
          if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
            Exit;
          end;
        end;

        //对该工作表进行全部替换
        SheetReplaceAllCount := ReplaceAll(FindReplaceRange);
        if ReplaceAllCount < 0 then begin
          break;
        end
        else begin
          ReplaceAllCount := ReplaceAllCount + SheetReplaceAllCount;
        end;
      end
      else
        Exit;
    end;
  end
  else begin
    //如果当前只选中了一个工作表
    if IsSingleCell(FindReplaceRange) then begin
      //如果选中的是一个单元格，需要把整表作为待查范围
      //如果选中的是两个以上单元格，当前的选择区域就是待查范围，不必额外处理
      ActiveSheet := FApplication.ActiveSheet;
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end
    end;

    //进行全部替换
    ReplaceAllCount := ReplaceAll(FindReplaceRange);
  end;

  if ReplaceAllCount = 0 then begin
    KSO_MessageBox(WideLoadResString(@sFindReplace_CannotReplace), '', MB_OK + MB_ICONEXCLAMATION);
    mcbFind.SetFocus;
  end
  else if ReplaceAllCount > 0 then begin
    KSO_MessageBox(WideFormat{TNT}(WideLoadResString(@sFound), [ReplaceAllCount]), '', MB_OK + MB_ICONEXCLAMATION);
    btnReplaceAll.SetFocus;
  end
  else if ReplaceAllCount < 0 then begin
      //替换过程中发现问题
    if ReplaceAllCount <> -2 then  // 和内核约定表示什么都没做
      KSO_MessageBox(WideLoadResString(@sReplace_ReplaceFailed), '', MB_OK + MB_ICONEXCLAMATION);
    btnReplaceAll.SetFocus;
  end;
end;

procedure TFindReplaceDlg.OnEnterSearch(Sender: TObject);
begin
  // ET组反映半透明影响性能，故注释。2004-7-11 -tifi
//  AlphaBlend := True;
//  AlphaBlendValue := 200;
end;

procedure TFindReplaceDlg.OnExitSearch(Sender: TObject);
begin
  // ET组反映半透明影响性能，故注释。2004-7-11 -tifi
//  AlphaBlend := False;
//  AlphaBlendValue := 255;
end;

function TFindReplaceDlg.GetFindLookin: TOleEnum;
const
  vFindLookin: array[TSearchLookin] of TOleEnum =
  (Integer(etFormulas), Integer(etValues), Integer(etComments));
begin
  result := vFindLookin[SearchLookin];
end;

function TFindReplaceDlg.GetFindLookAt: TOleEnum;
const
  vFindLookAt: array[Boolean] of TOleEnum = (etPart, etWhole);
begin
  result := vFindLookAt[cbMatchContents.Checked];
end;

function TFindReplaceDlg.GetSearchOrder: TOleEnum;
const
  vSearchOrder: array[TSearchOrder] of TOleEnum = (etByRows, etByColumns);
begin
  result := vSearchOrder[SearchOrder];
end;

function TFindReplaceDlg.GetMatchCase: OleVariant;
begin
  result := cbMatchCase.Checked;
end;

function TFindReplaceDlg.GetMatchByte: OleVariant;
begin
  result := cbMatchByte.Checked;
end;


function TFindReplaceDlg.Find(FRange: IETRange): IETRange;
begin
  //查找
  if not _FindNextFlag then begin
    Result := FRange.Find(WideString(mcbFind.Text), FApplication.ActiveCell, GetFindLookin,
      GetFindLookAt, GetSearchOrder, etNext, GetMatchCase, GetMatchByte, NULL);
    _FindNextFlag := True;
    _FindCompleteFlag := false;
  end
  else if _FindCompleteFlag then begin
    Result := nil;
  end
  else begin
    Result := FRange.FindNext(FApplication.ActiveCell);
  end;
end;

function TFindReplaceDlg.Replace(FRange: IETRange; var fIsReplaceFail: Boolean): IETRange;
var
  ReplaceCount: Integer;
  FindRes: IETRange;
begin
   //对于单表替换
   //总是替换当前单元格，然后找到下一个待替换的单元格，设置为当前单元格
   //这样的话，如果当前单元格不能替换成功，则按一下“替换”的结果是尝试找到第一个代替换内容
  ReplaceCount := FApplication.ActiveCell.Replace(WideString(mcbFind.Text), WideString(mcbReplace.Text), GetFindLookAt,
    GetSearchOrder, GetMatchCase, GetMatchByte, NULL, NULL);

  if ReplaceCount > 0 then begin
      //如果当前单元格替换成功
      //找下一个
    if _ReplaceFirstFlag then begin
       //如果是第一次替换，对于后续的用Find
      FindRes := FRange.Find(WideString(mcbFind.Text), FApplication.ActiveCell, GetFindLookin,
        GetFindLookAt, GetSearchOrder, etNext, GetMatchCase, GetMatchByte, NULL);
      _ReplaceFirstFlag := false;
    end
    else begin
       //否则用FindNext
      FindRes := FRange.FindNext(FApplication.ActiveCell);
    end;

      //确定返回什么
    if FindRes <> nil then begin
        //若能找到下一个，则返回下一个
      Result := FindRes;
    end
    else begin
        //反之返回当前单元格
      Result := FApplication.ActiveCell;
    end
  end
  else if ReplaceCount = 0 then begin
      //如果当前单元格不匹配，因此没有替换次数
      //直接找下一个返回
    if _ReplaceFirstFlag then begin
       //如果是第一次替换，对于后续的用Find
      Result := FRange.Find(WideString(mcbFind.Text), FApplication.ActiveCell, GetFindLookin,
        GetFindLookAt, GetSearchOrder, etNext, GetMatchCase, GetMatchByte, NULL);
      _ReplaceFirstFlag := false;
    end
    else begin
       //否则用FindNext
      Result := FRange.FindNext(FApplication.ActiveCell);
    end;
  end
  else if ReplaceCount = -2 then begin
    // 如果是替换失败，即要填入的内容可能不是合法的公式，或者出于保护状态
    // -2是和Core约定的值
    FbProtected := True;
    Result := nil;
    end
  else if ReplaceCount = -1 then begin
    Result := nil;
    fIsReplaceFail := True;
    end
  else begin
    // 其他情况不弹框
    Result := nil;
    fIsReplaceFail := False;
    end;
end;

function TFindReplaceDlg.ReplaceAll(FRange: IETRange): Integer;
begin
  Result := FindReplaceRange.Replace(WideString(mcbFind.Text), WideString(mcbReplace.Text), GetFindLookAt,
    GetSearchOrder, GetMatchCase, GetMatchByte, NULL, NULL);
end;

function TFindReplaceDlg.IsSingleCell(FRange: IETRange): Boolean;
var
  FirstItem: IDispatch; //在参数区域count不为1时使用
  FirstCell: IETRange;
  MergeArea: IETRange;
  MergeAreaAddress: WideString; //经过合并判断的区域的地址
  RangeAddress: WideString; //原始参数区域的地址
  RelativeTo: OleVariant;
begin
  //如果本身区域的count为1，则肯定是单个单元格
  //如果该区域count不为1，则要判断是否整个为一个合并单元格
  //方法是：
  //从该区域中取一个单元格。可以使用item(1,1)
  //对该单元格求 MergeArea ，然后看所得区域是否和原区域一致，如果一致，则说明整个区域是个合并单元格
  Result := false;

  if FRange.Count = 1 then begin
    Result := true;
  end
  else begin
    //如果该区域count不为1
    FirstItem := FRange.item[1, 1];
    if FirstItem.QueryInterface(IETRange, FirstCell) = S_OK then begin
      MergeArea := FirstCell.MergeArea;
      //这里面参数给得比较随意，以后要改（目前这几个参数API实际上也没有处理）
      MergeAreaAddress := MergeArea.Address[false, false, etA1, false, RelativeTo];
      RangeAddress := FRange.Address[false, false, etA1, false, RelativeTo];

      if MergeAreaAddress = RangeAddress then begin
        //如果合并单元格的地址等于参数地址
        //则说明参数是个合并单元格
        Result := true;
      end;
    end;
  end;
end;

procedure TFindReplaceDlg.FormActivate(Sender: TObject);
var
  bCtrlEnabled: Boolean;
  I: Integer;
begin
  bCtrlEnabled := (FApplication.Workbooks.Count <> 0);
  btnReplaceAll.Enabled := bCtrlEnabled;
  btnReplace.Enabled := bCtrlEnabled;
  btnFindNext.Enabled := bCtrlEnabled;
  btnFindAll.Enabled := bCtrlEnabled;
  if (FApplication <> nil) and (lvFindAll.Items.Count <> 0) then
  begin
    for I := FApplication.Workbooks.Count downto 1 do
    begin
      if FApplication.Workbooks.Item[I].Name = lvFindAll.Items[0].Caption then
        Exit;
    end;
    lvFindAll.Clear;
  end;
end;

class function TFindReplaceDlg.SD_GetType_: TDialogType;
begin
  result := ID_FindReplace;
end;

procedure TFindReplaceDlg.SD_ObjParamChange;
begin
  inherited;
  if (FObjParam = nil) or
    (FObjParam.QueryInterface(IETApplication, FApplication) <> S_OK) then
    FApplication := nil;
end;

procedure TFindReplaceDlg.SD_EnvParamChange;
var
  FDlgMode: IFindReplaceDialogMode;
begin
  inherited;
  if (FEnvParam <> nil) then begin
    if Supports(FEnvParam, IFindReplaceDialogMode, FDlgMode) then begin
      FindReplaceMode := TFindReplaceMode(FDlgMode.FindReplaceDialogMode);
    end
    else begin
      Assert(False);
    end;
  end;
end;

function TFindReplaceDlg.SD_Update: HResult;
begin
  result := S_OK;
end;

procedure TFindReplaceDlg.CreateWnd;
begin
  inherited CreateWnd;
  SendMessage(Handle, WM_SETICON, ICON_SMALL, 0);
  SendMessage(Handle, WM_SETICON, ICON_BIG, 0);
  SendMessage(Handle, WM_SETICON, ICON_SMALL2, 0);
end;

procedure TFindReplaceDlg.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.ExStyle := Params.ExStyle or WS_EX_DLGMODALFRAME;
end;

procedure TFindReplaceDlg.DoShow;
begin
  inherited;

  ClientHeight := ToCurPixelsPerInchValue(iHeightAdvancedMode);
  ClientWidth := ToCurPixelsPerInchValue(iWidthAnyMode);


  if FindReplaceMode = frmGoto then
  begin
    FCanShowCombo := True;
    SwitchMode(tcGoto);
  end else
  begin
    FCanShowCombo := True;
    SwitchMode(tcFindReplace);
  end;
end;

constructor TFindReplaceDlg.Create(AOwner: TComponent);
begin
  DialogMode := dmNormal;
  LastDialogMode := dmAdvanced;

  if (AOwner is TFindReplaceDlg) then
    AOwner := Application.MainForm;
  inherited Create(AOwner);

  FormStyle := fsNormal;

  mcbFind.MRUList := dmMain.MRUStringListFind;
  mcbReplace.MRUList := dmMain.MRUStringListReplace;

  _FindNextFlag := False;

  _ReplaceFirstFlag := True;
  _ReplaceSecondFlag := False;
  _ReplaceSinceFlag := False;
  _ReplaceCompleteFlag := False;
  _FirstFindInSheet := False;

  FSwitchFindReplaceMode[frmFind] := SwitchToFind;
  FSwitchFindReplaceMode[frmReplace] := SwitchToReplace;
  FSwitchFindReplaceMode[frmGoto] := SwitchToGoto;
  FSwitchDialogMode[dmNormal] := SwitchToNormal;
  FSwitchDialogMode[dmAdvanced] := SwitchToAdvanced;
  FbProtected := False;

  RestoreState;
  cbMatchByte.Visible:= BOOL(KsoQueryFeatureState(kaf_kso_FullWidthHalfWidth));
  OnActivate := FormActivate;
  ChangeTooBarShortCut(Self, True);

  btnReplaceAll.Left := tcFindReplace.Left;
  btnReplace.Left := btnReplaceAll.Left + btnReplaceAll.Width + 6;
  btnFindAll.Left := btnReplace.Left + btnReplace.Width + 6;
end;

procedure TFindReplaceDlg.SaveState;
begin
  FindContent := mcbFind.Text;
  ReplaceContent := mcbReplace.Text;
end;

procedure TFindReplaceDlg.RestoreState;
var
  OrgSearchLookin: TSearchLookin;
begin
  mcbFind.ItemIndex := 0;
  mcbReplace.ItemIndex := 0;

  OrgSearchLookin := SearchLookin;
  FSwitchFindReplaceMode[FindReplaceMode];
  SearchLookin := OrgSearchLookin;
  // FSwitchDialogMode[DialogMode];

  cbRange.ItemIndex := Integer(SearchRange);
  cbOrder.ItemIndex := Integer(SearchOrder);
  cbLookin.ItemIndex := Integer(SearchLookin);

  if soMatchCase in SearchOptions then
    cbMatchCase.Checked := True;
  if soMatchContents in SearchOptions then
    cbMatchContents.Checked := True;
  if (soMatchByte in SearchOptions) and (BOOL(KsoQueryFeatureState(kaf_kso_FullWidthHalfWidth))) then
    cbMatchByte.Checked := True;

  mcbFind.Text := FindContent;
  mcbReplace.Text := ReplaceContent;
   
  SetCellType(CellType);
  SetDataType(DataType);
end;

procedure TFindReplaceDlg.SwitchToFind;
begin
  tcFindReplace.TabIndex := 0;

  tcGoto.Hide;
  btnGoto.Hide;
  tcFindReplace.Show;

  lblReplace.Hide;
  mcbReplace.Hide;
  btnReplaceAll.Hide;
  btnReplace.Hide;
  btnFindNext.Show;
  btnFindAll.Show;
  if FCanShowCombo then
    mcbFind.SetFocus;

  with cbLookin.Items do begin
    BeginUpdate;
    Clear;
    Add(WideLoadResString(@sLookinFormulas));
    Add(WideLoadResString(@sLookinValues));
    Add(WideLoadResString(@sLookinComments));
    EndUpdate;
  end;

  cbLookin.ItemIndex := Integer(FFindSearchLookin);
  FFindSearchLookin := SearchLookin;
  cbLookin.OnSelect(Self);
end;

procedure TFindReplaceDlg.SwitchToReplace;
begin
  tcFindReplace.TabIndex := 1;

  tcGoto.Hide;
  btnGoto.Hide;
  tcFindReplace.Show;

  lblReplace.Show;
  mcbReplace.Show;
  btnReplaceAll.Show;
  btnReplace.Show;
  btnFindNext.Show;
  btnFindAll.Show;
  if FCanShowCombo then
    if mcbFind.Text = EmptyStr then begin
      mcbReplace.SelLength := 0;
      mcbFind.SetFocus;
    end
    else begin
      mcbFind.SelLength := 0;
      mcbReplace.SetFocus;
    end;

  FFindSearchLookin := TSearchLookin(cbLookin.ItemIndex);
  with cbLookin.Items do begin
    BeginUpdate;
    Clear;
    Add(WideLoadResString(@sLookinFormulas));
    EndUpdate;
  end;
  cbLookin.ItemIndex := 0;
  cbLookin.OnSelect(Self);

  SearchLookin := slFormulas;
end;

procedure TFindReplaceDlg.SwitchToGoto;
begin
  tcGoto.TabIndex := 2;

  tcGoto.Show;
  btnGoto.Show;
  tcFindReplace.Hide;

  btnReplaceAll.Hide;
  btnReplace.Hide;
  btnFindAll.Hide;
  btnFindNext.Hide;
  GoToObject:= rbObject.Checked;
  if CellType = 0 then
  begin
    rbData.OnClick(rbData);
    rbData.SetFocus;
  end;
end;

procedure TFindReplaceDlg.SwitchToNormal;
var
  HeightChanged: Integer;
begin
  if LastDialogMode = dmNormal then
    Exit;

  Constraints.MaxHeight := 0;
  Constraints.MinHeight := 0;
  Constraints.MaxWidth := 0;
  Constraints.MinWidth := 0;

  HeightChanged := ToCurPixelsPerInchValue(iHeightNormalMode) - ToCurPixelsPerInchValue(iHeightAdvancedMode);
  ClientHeight := ClientHeight + HeightChanged;
  ClientWidth := ToCurPixelsPerInchValue(iWidthAnyMode);

  btnClose.Top := btnClose.Top + HeightChanged;
  btnFindAll.Top := btnFindAll.Top + HeightChanged;
  btnFindNext.Top := btnFindNext.Top + HeightChanged;
  btnGoto.Top := btnGoto.Top + HeightChanged;
  btnReplace.Top := btnReplace.Top + HeightChanged;
  btnReplaceAll.Top := btnReplaceAll.Top + HeightChanged;
  lvFindAll.Top := lvFindAll.Top + HeightChanged;
  lvFindAll.Height := MAX(ClientHeight - lvFindAll.Top - ScaleUISizeY(8), 0);

  tcFindReplace.Height := tcFindReplace.Height + HeightChanged;

  Constraints.MinHeight := ToCurPixelsPerInchValue(iHeightNormalMode) + Height - ClientHeight;
  Constraints.MinWidth := Width;

  lblWithin.Hide;
  lblSearch.Hide;
  lblLookin.Hide;
  cbRange.Hide;
  cbOrder.Hide;
  cbLookin.Hide;
  cbMatchCase.Hide;
  cbMatchContents.Hide;
  if BOOL(KsoQueryFeatureState(kaf_kso_FullWidthHalfWidth)) then
    cbMatchByte.Hide;

  btnOptions.Caption := WideLoadResString(@sNormalModeCap);
  LastDialogMode := dmNormal;
end;

procedure TFindReplaceDlg.SwitchToAdvanced;
var
  HeightChanged: Integer;
begin
  if LastDialogMode = dmAdvanced then
    Exit;

  Constraints.MaxHeight := 0;
  Constraints.MinHeight := 0;
  Constraints.MaxWidth := 0;
  Constraints.MinWidth := 0;

  HeightChanged := ToCurPixelsPerInchValue(iHeightAdvancedMode) - ToCurPixelsPerInchValue(iHeightNormalMode);
  ClientHeight := ClientHeight + HeightChanged;
  ClientWidth := ToCurPixelsPerInchValue(iWidthAnyMode);

  btnClose.Top := btnClose.Top + HeightChanged;
  btnFindAll.Top := btnFindAll.Top + HeightChanged;
  btnFindNext.Top := btnFindNext.Top + HeightChanged;
  btnGoto.Top := btnGoto.Top + HeightChanged;
  btnReplace.Top := btnReplace.Top + HeightChanged;
  btnReplaceAll.Top := btnReplaceAll.Top + HeightChanged;

  lvFindAll.Top := lvFindAll.Top + HeightChanged;
  lvFindAll.Height := MAX(ClientHeight - lvFindAll.Top - ScaleUISizeY(8), 0);

  tcFindReplace.Height := tcFindReplace.Height + HeightChanged;


  Constraints.MinHeight := ToCurPixelsPerInchValue(iHeightAdvancedMode) + Height - ClientHeight;
  Constraints.MinWidth := Width;

  if FindReplaceMode = frmGoto then
  begin
  
  end else
  begin
    lblWithin.Show;
    lblSearch.Show;
    lblLookin.Show;
    cbRange.Show;
    cbOrder.Show;
    cbLookin.Show;
    cbMatchCase.Show;
    cbMatchContents.Show;
    if BOOL(KsoQueryFeatureState(kaf_kso_FullWidthHalfWidth)) then
      cbMatchByte.Show;

    btnOptions.Caption := WideLoadResString(@sAdvancedModeCap);
  end;

  LastDialogMode := dmAdvanced;
end;

procedure TFindReplaceDlg.DoClose(var Action: TCloseAction);
begin
  inherited;
  SaveState;
end;

// -----------------------------------------------------------------------

destructor TFindReplaceDlg.Destroy;
begin
  ChangeTooBarShortCut(Self, False);
  if Assigned(DialogEvent) then
    DialogEvent := nil;

  inherited;
end;

procedure TFindReplaceDlg.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (ssCtrl in Shift) and (Key = VK_TAB) then
  begin
    if not (ssShift in Shift) then
    begin
      if tcFindReplace.TabIndex = 0 then
        tcFindReplace.TabIndex := 1
      else if tcFindReplace.TabIndex = 1 then
        tcFindReplace.TabIndex := 2
      else
        tcFindReplace.TabIndex := 0;
    end else
    begin
      if tcFindReplace.TabIndex = 0 then
        tcFindReplace.TabIndex := 2
      else if tcFindReplace.TabIndex = 1 then
        tcFindReplace.TabIndex := 0
      else
        tcFindReplace.TabIndex := 1;
    end;

    tcFindReplace.OnChange(tcFindReplace);
  end;
end;

procedure TFindReplaceDlg.EnableDataControl(const bEnabled: Boolean);
begin
  cbConstants.Enabled := bEnabled;
  cbFormulas.Enabled := bEnabled;
  cbNumbers.Enabled := bEnabled;
  cbText.Enabled := bEnabled;
  cbLogicals.Enabled := bEnabled;
  cbErrors.Enabled := bEnabled;
  lblDataType.Enabled := bEnabled;
  if DataValue <> 0 then
  begin
    if DataValue = etCellTypeConstants then
      SetState_Whist(cbConstants, cbChecked)
    else if DataValue = etCellTypeFormulas then
      SetState_Whist(cbFormulas, cbChecked)
    else
    begin
      SetState_Whist(cbConstants, cbChecked);
      SetState_Whist(cbFormulas, cbChecked);
    end;
  end;
  if DataType <> 0 then
  begin
    SetDataType(DataType);
  end;
end;

function TFindReplaceDlg.GetDataType: ETSpecialCellsValue;
begin
  Result := $00000000;
  if cbNumbers.Checked then
    Inc(Result, etNumbers);

  if cbText.Checked then
    Inc(Result, etTextValues);

  if cbLogicals.Checked then
    Inc(Result, etLogical);

  if cbErrors.Checked then
    Inc(Result, etErrors);

  if Result = 0 then
    Result := etNumbers + etTextValues + etLogical + etErrors;
end;

function TFindReplaceDlg.GetCellType: ETCellType;
begin
  if (cbFormulas.Checked) and (cbConstants.Checked) then
    Result := etCellFormulasAndConstants
  else if cbConstants.Checked then
    Result := etCellTypeConstants
  else if (cbFormulas.Checked) then
    Result := etCellTypeFormulas
  else
    Result := etCellFormulasAndConstants;
end;

procedure TFindReplaceDlg.rbDataClick(Sender: TObject);
var
  component: TComponent;
begin
  GoToObject:= False;
  component := Sender as TComponent;
  case component.Tag of
    0:
    begin
      CellType := GetCellType;
      DataType := GetDataType;
    end;
    1:
    begin
      CellType := etCellTypeComments;
    end;
    2:
    begin
      CellType := etCellTypeBlanks;
    end;
    3:
    begin
      CellType := etCellTypeVisible;
    end;
    4:
    begin
      CellType := etCellTypeLastCell;
    end;
    5:
    begin
      CellType := etCellCurDataArea;
    end;
    6:
    begin
      CellType := etCellHyperlink;
    end;
    7:
    begin
      GoToObject := True; 
    end;
  end;
  DataValue := GetCellType;
  EnableDataControl(rbData.Checked);
end;

procedure TFindReplaceDlg.btnGotoClick(Sender: TObject);
var
  Selection: IDispatch;
  LocateRes: IETRange;
  ActiveSheet: IDispatch;
  OvActiveSheet: OleVariant;
  WorkSheet: IWorkSheet;
  bProtect: WordBool;
begin
  if FApplication = nil then
    Exit;

  ActiveSheet := FApplication.ActiveSheet;
  if ActiveSheet = nil then
    Exit;

  if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then
  begin
    bProtect := WorkSheet.ProtectContents;
    if bProtect then
    begin
      // 提示不可定位信息
      KSO_MessageBox(WideLoadResString(@sCannotGoto), '', MB_ICONEXCLAMATION + MB_OK);
      Exit;
    end;
  end else
    Exit;

  if GoToObject then begin
    OvActiveSheet:= FApplication.ActiveSheet;
    Assert(not VarIsEmpty(OvActiveSheet.Shapes));
    if OvActiveSheet.Shapes.Count = 0 then
      KSO_MessageBox(WideLoadResString(@sLocate_NotFindObject), '', MB_ICONEXCLAMATION + MB_OK)
    else
      OvActiveSheet.Shapes.SelectAll;
    end
  else begin
    Selection := FApplication.Selection[$0804];
    if (Selection = nil) then
      Exit;

    if Selection.QueryInterface(IETRange, FLocateRange) <> S_OK then
      FLocateRange := WorkSheet.Range['A1', EmptyParam];

    if FLocateRange = nil then
      Exit;

    LocateRes := FLocateRange.SpecialCells(CellType, DataType);

    if LocateRes <> nil then
      LocateRes.Select
    else
      KSO_MessageBox(WideLoadResString(@sLocate_NotFind), '', MB_ICONEXCLAMATION + MB_OK);
    end;  
end;

procedure TFindReplaceDlg.cbNumbersClick(Sender: TObject);
begin
  if not cbNumbers.Checked and not cbText.Checked and not cbLogicals.Checked and not cbErrors.Checked then
  begin
    SetChecked_Whist(cbNumbers, true);
    SetChecked_Whist(cbText, true);
    SetChecked_Whist(cbLogicals, true);
    SetChecked_Whist(cbErrors, true);
  end;
  DataType := GetDataType;
end;

procedure TFindReplaceDlg.cbConstantsClick(Sender: TObject);
begin
  if not cbConstants.Checked and not cbFormulas.Checked then
  begin
    SetChecked_Whist(cbConstants, true);
    SetChecked_Whist(cbFormulas, true);
  end;
  CellType := GetCellType;
end;

procedure TFindReplaceDlg.FormCreate(Sender: TObject);
begin
  // Caption := WideStripHotkey(tcFindReplace.Tabs[tcFindReplace.TabIndex]);
  // 暂时屏蔽掉超链接功能
//  rbHyperlink.Visible := false;
end;

procedure TFindReplaceDlg.SetCellType(const Value: ETCellType);
begin
  case Value of
    etCellFormulasAndConstants:
    begin
      SetChecked_Whist(rbData, True);
      SetState_Whist(cbConstants, cbChecked);
      SetState_Whist(cbFormulas, cbChecked);
      SetDataType(DataType);
    end;
    etCellTypeFormulas:
    begin
      SetChecked_Whist(rbData, True);
      SetState_Whist(cbFormulas, cbChecked);
      SetDataType(DataType);
    end;
    etCellTypeConstants:
    begin
      SetChecked_Whist(rbData, True);
      SetState_Whist(cbConstants, cbChecked);
      SetDataType(DataType);
    end;
    etCellTypeComments:
    begin
      SetChecked_Whist(rbComments, True);
      SetDataType(DataType);
      EnableDataControl(rbData.Checked);
    end;
    etCellTypeBlanks:
    begin
      SetChecked_Whist(rbBlanks, True);
      SetDataType(DataType);
      EnableDataControl(rbData.Checked);
    end;
    etCellTypeVisible:
    begin
      SetChecked_Whist(rbVisibleCellsOnly, True);
      SetDataType(DataType);
      EnableDataControl(rbData.Checked);
    end;
    etCellTypeLastCell:
    begin
      SetChecked_Whist(rbLastCell, True);
      SetDataType(DataType);
      EnableDataControl(rbData.Checked);
    end;
    etCellHyperlink:
    begin
      SetChecked_Whist(rbHyperlink, True);
      SetDataType(DataType);
      EnableDataControl(rbData.Checked);
    end;
    etCellCurDataArea:
    begin
      SetChecked_Whist(rbCurDataArea, True);
      SetDataType(DataType);
      EnableDataControl(rbData.Checked);
    end;
  end;
end;

procedure TFindReplaceDlg.SetDataType(const Value: ETSpecialCellsValue);
begin
  if LongBool(Value and etNumbers) then
    SetState_Whist(cbNumbers, cbChecked);

  if LongBool(Value and etTextValues) then
    SetState_Whist(cbText, cbChecked);

  if LongBool(Value and etLogical) then
    SetState_Whist(cbLogicals, cbChecked);

  if LongBool(Value and etErrors) then
    SetState_Whist(cbErrors, cbChecked);
end;

function TFindReplaceDlg.IsShortCut(var Message: TWMKey): Boolean;
var
  shift: TShiftState;
begin
  shift:= KeyDataToShiftState(Message.KeyData);
  if (Message.CharCode = VK_F4) and (ssCtrl in shift) then begin
    mcbFind.DroppedDown:= True;
    Result:= True;
    end
  else
    Result := inherited IsShortCut(Message);
end;


procedure TFindReplaceDlg.OnSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
  ActiveSheet: IDispatch;
  Selection: IDispatch;
  ActiveWorkbook: IDispatch; //遍历中使用
  WorkSheet: IWorkSheet;
  WorkBook: IWorkbook;
  AddressStr: WideString;
  I: Integer;
begin
  if(lvFindAll.Items.Count = 0) then Exit;
  if(lvFindAll.Selected = nil) then begin
    WorkSheetStr := '';
    Exit;
  end;
  Selection := FApplication.Selection[$0804];
  I := lvFindAll.Selected.Index;
  if (Selection = nil) then exit;
  if FApplication = nil then
    Exit;
  ActiveWorkbook := FApplication.Workbooks[lvFindAll.Items[0].Caption];
  if ActiveWorkbook.QueryInterface(IWorkbook, Workbook) = S_OK then begin
    WorkBook.Activate;
  end else
    Exit;
  if (SearchRange = srWorkBook) and (WorkSheetStr = '') then
    WorkSheetStr := lvFindAll.Selected.SubItems.Strings[0];

  for I := 0 to lvFindAll.Items.Count - 1 do begin
    if lvFindAll.Items[I].Selected = True then begin
      if AddressStr = '' then begin
        if (SearchRange = srWorkBook) and (WorkSheetStr <> lvFindAll.Items[I].SubItems.Strings[0]) then
          lvFindAll.Items[I].Selected := False;
        AddressStr := AddressStr + lvFindAll.Items[I].SubItems.Strings[2];
        Continue;
      end else begin
        if (SearchRange = srWorkBook) and (WorkSheetStr <> lvFindAll.Items[I].SubItems.Strings[0])then
          lvFindAll.Items[I].Selected := False;
        AddressStr := AddressStr + ',' +lvFindAll.Items[I].SubItems.Strings[2];
        Continue;
      end;
    end;
  end;

  ActiveSheet := FApplication.ActiveWorkbook.Worksheets[lvFindAll.Selected.SubItems.Strings[0]];
  if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
    WorkSheet.Activate;
    if AddressStr <> '' then
      WorkSheet.Range[AddressStr,EmptyParam].Select;
  end;

  {AddressStr := '';
  if (lvFindAll.Selected = nil) or (lvFindAll.Items.Count = 0) then Exit;
  Selection := FApplication.Selection[$0804];
  I := lvFindAll.Selected.Index;
  if (Selection = nil) then exit;
  if FApplication = nil then
    Exit;
  ActiveWorkbook := FApplication.Workbooks[lvFindAll.Items[0].Caption];
  if ActiveWorkbook.QueryInterface(IWorkbook, Workbook) = S_OK then begin
    WorkBook.Activate;
  end else
    Exit;


//  while lvFindAll.Selected.Index <> I do begin
  repeat
    Inc(I);
    I := I mod lvFindAll.Items.Count;
    if lvFindAll.Items[I].Selected = True then begin
      if AddressStr = '' then begin
          AddressStr := AddressStr + lvFindAll.Items[I].SubItems.Strings[2];
        if (lvFindAll.Selected.SubItems.Strings[0] <> lvFindAll.Items[I].SubItems.Strings[0]) then begin
          lvFindAll.Items[lvFindAll.Selected.Index].Selected := False;
        end;
      end else begin
        AddressStr := AddressStr + ',' +lvFindAll.Items[I].SubItems.Strings[2];
      end;
    end;

//  end;
  until lvFindAll.Selected.Index = I;



//  for I := lvFindAll.Selected.Index + 1 to lvFindAll.Selected.Index do begin
//    I := I mod lvFindAll.Items.Count;
//    if lvFindAll.Items[I].Selected = True then begin
//      if AddressStr = '' then begin
//        AddressStr := AddressStr + lvFindAll.Items[I].SubItems.Strings[2];
//        if lvFindAll.Selected.SubItems.Strings[0] <> lvFindAll.Items[I].SubItems.Strings[0] then begin
//          lvFindAll.Items[lvFindAll.Selected.Index].Selected := False;
//        end;
//        Continue;
//      end else begin
//        AddressStr := AddressStr + ',' +lvFindAll.Items[I].SubItems.Strings[2];
//        Continue;
//      end;
//    end;
//  end;

  ActiveSheet := FApplication.ActiveWorkbook.Worksheets[lvFindAll.Selected.SubItems.Strings[0]];
  if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
    WorkSheet.Activate;
//    WorkSheet.Range[lvFindAll.Selected.SubItems.Strings[2], lvFindAll.Selected.SubItems.Strings[2]].Select;

    if AddressStr <> '' then
      WorkSheet.Range[AddressStr,EmptyParam].Select;
  end; }

end;

procedure TFindReplaceDlg.OnColClick(Sender: TObject; Column: TListColumn);
begin
  //暂时不增加
  ColumnToSort := Column.Index;
  lvFindAll.AlphaSort;
  asc := not asc;
end;

procedure TFindReplaceDlg.OnClick(Sender: TObject);
begin
  //暂时不增加
end;


procedure TFindReplaceDlg.lvFindAllResize(Sender: TObject);
begin
  lvFindAll.Columns[lvFindAll.Columns.Count - 1].Width := lvFindAll.ClientWidth -
            colSetWidth(lvFindAll.Columns.Count - 1);
end;

procedure TFindReplaceDlg.lvFindAllCompare(Sender: TObject; Item1,
  Item2: TListItem; Data: Integer; var Compare: Integer);
var 
  ix: Integer; 
begin
  if ColumnToSort = 0 then
    Compare := CompareText(Item1.Caption,Item2.Caption) 
  else begin 
   ix := ColumnToSort - 1; 
   Compare := CompareText(Item1.SubItems[ix],Item2.SubItems[ix]); 
  end;
  if asc then
    Compare := -Compare;
end;


initialization
  RegisterDialogClass(TFindReplaceDlg.SD_GetType_, TFindReplaceDlg);

finalization
  UnregisterDialogClass(TFindReplaceDlg.SD_GetType_);

end.
