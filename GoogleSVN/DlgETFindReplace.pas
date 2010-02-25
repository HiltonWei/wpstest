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
 #C1:2007-03-09:����bug:26096

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
    _ReplaceFirstFlag: Boolean; //��һ���滻
    _ReplaceSecondFlag: Boolean; //�ڶ����滻
    _ReplaceSinceFlag: Boolean; //�Ժ���滻
    _ReplaceCompleteFlag: Boolean; //�滻���
    _FindCompleteFlag: Boolean; //�������
    _FirstFindInSheet: Boolean;
    FindFirst: IETRange;//����ȫ���ĵ�һ��

    SheetsCount: Integer; //�������й���������

    //FRange: IETRange;
    FindReplaceRange: IETRange;
    FLocateRange: IETRange;

    FirstFindRangeRow: integer; //��ĳһ���������ҵ��ĵ�һ����Ԫ����кţ��������ʵ�ʱ����ֹ��ĳ���еĲ���
    FirstFindRangeCol: integer; //��ĳһ���������ҵ��ĵ�һ����Ԫ����кţ��������ʵ�ʱ����ֹ��ĳ���еĲ���

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

    function IsSingleCell(FRange: IETRange): Boolean; //������Ԫ�񡢺ϲ���Ԫ�����ǵ�����Ԫ��

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
  //�����ֶԻ���ʧȥ���㣬����һ����Ϊ���¿�ʼ�µĲ����滻
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
  NextSheet: IDispatch; //������ʹ��
  WorkSheet: IWorkSheet;
  FindRes: IETRange;
  NoFindOutSheetCount: Integer; //���ҹ���û���ҵ��������ݵĹ�������Ŀ�����ڹ������Ͷ�ѡ������ʱ�Ĳ���
  SheetIndex: Integer;
  SelSheets: Sheets;
begin
  // todo: �������ȥ�����棬Ӧ����ϸ����߼���ϵ��
  SheetIndex := 0;

  if FApplication = nil then
    Exit;
  if Assigned(FApplication) then
    SelSheets := FApplication.ActiveWindow.SelectedSheets;

  Selection := FApplication.Selection[$0804];
  if (Selection = nil) then exit;

  //�������������Ϊ�գ���ʾ����
  if mcbFind.Text = '' then begin
    KSO_MessageBox(WideLoadResString(@sFindReplace_InvalidWhat), '', MB_OK + MB_ICONEXCLAMATION);
    mcbFind.SetFocus;
    Exit;
  end;

  { add by oyxz, see #C1 }
  // �������ƹ����Ĵ��룬���Ϊ��ֵʱ�����������²���������ᵼ��bug:26096
  mcbFind.ActivateCurrItem;
  mcbReplace.ActivateCurrItem;
  { add end }

  //�����ǰ��ѡ����������ʲôҲ����
  if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
    Exit;
  end;

  //ȷ�����ҷ�Χ��ע���е�������ظı�
  //���ڹ�������һ�β���֮�󶼲��ظı�
  if SearchRange = srWorkBook then begin
    //�����ǰ���ҷ�Χ�Ǹ�������

    //����ǵ�һ�β��ң����ȡ��ǰ��������Cells�������ҷ�Χ������Ҫ�����ǰ��Selection
    if not _FindNextFlag then begin
      SheetsCount := FApplication.ActiveWorkbook.Worksheets.Count;
      _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
    end;

    ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
    if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
      FindReplaceRange := WorkSheet.Cells;
       //FApplication.ActiveCell.Select;//ȡ��ԭ�������ϵ�����ѡ��
    end
    else begin
      Exit;
    end;

    //������һ�㣬Ĭ���õ�ǰ���ҷ�Χ���в��ң����ϴ����µģ�
    FindRes := Find(FindReplaceRange);

    //�����Ƿ񻻱����
    //����ҲҪ��ô��  ���л�������һ��������
    if FindRes <> nil then begin
      //����ҵ���
      if _FirstFindInSheet then begin
        //�����ڵ�ǰ�������е�һ�η���
        //�ѱ�־��Ϊ�٣������ҵ������ݱ�������
        _FirstFindInSheet := false;
        FirstFindRangeRow := FindRes.Row;
        FirstFindRangeCol := FindRes.Column;
      end
      else begin
        //������ǣ���Ƚ��ҵ��ĺͱ����
        if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
           //���һ�£�˵������һ���ˣ�Ҫ������һ�β��ң��Ѳ��ҽ������
          FindRes := nil;
        end;
      end;
    end;

    if FindRes = nil then begin
       //û�ҵ��������ݵĹ�������Ŀ��ʼ��Ϊ0
      NoFindOutSheetCount := 0;
       //��ȡ��ǰ����������Ϊ����������
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //û�ҵ��������ݹ�������Ŀ�Ѿ�Ϊ������������Ŀ
             //˵��������������û����Ҫ�����ݣ��˳�����
          break;
        end
        else begin
             //�ҵ���һ����������������������1ȡģ��
             //ģΪ������������1����Ϊ������������1��ʼ����ȡģ���Ϊ0��Ҫ����Ϊ1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := FApplication.ActiveWorkbook.Worksheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
            FindReplaceRange := WorkSheet.Cells;
               //���˹�����Ҫ����ѡ��״̬������ǰ���Ԫ������һ������
               //��Ӱ����ҵ���ʼλ��
            _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
          end;
        end;

        if WorkSheet.Visible <> etSheetVisible then
          FindRes := nil // ����������ɼ����򲻲���
        else
          FindRes := Find(FindReplaceRange);

          //�ж��Ƿ��һ���������һ��
        if FindRes <> nil then begin
            //����ҵ���
            //ֻ�е��ù�������ƥ������ʱ���Ž��ñ�ѡ��
          WorkSheet.Select(true);
            //���������ԭ��
            //1�����һ��ʼ��ѡ�й�������ù�������û�б��������ݾͻ������˸���ñ�һ��������
            //2�������ʼ��ѡ�й�������ң���ǰ��Ԫ�񽫻����ϸ�������ģ������±����λ��������
            //���Բ��õķ��������ң�����ҵ�����ѡ������
          Find(FindReplaceRange);
          if _FirstFindInSheet then begin
              //�����ڵ�ǰ�������е�һ�η���
              //�ѱ�־��Ϊ�٣������ҵ������ݱ�������
            _FirstFindInSheet := false;
            FirstFindRangeRow := FindRes.Row;
            FirstFindRangeCol := FindRes.Column;
          end
          else begin
              //������ǣ���Ƚ��ҵ��ĺͱ����
            if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
                 //���һ�£�˵������һ���ˣ�Ҫ������һ�β��ң��Ѳ��ҽ������
              FindRes := nil;
            end;
          end;
        end;

        if FindRes = nil then
             //��������ѡȡ��һ��������û���ҵ����򽫱�������Ŀ��1����ʾ�ñ�û�д���������
             //ѭ����Ĳ���ʧ�ܲ���֤�������޲������ݣ��п����ǲ鵽β��
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
    //�����ǰѡ���˶��������
    if not _FindNextFlag then begin
      _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
      //�˴���¼�Ĺ�������Ŀ�ǵ�ǰѡ�еĹ�������Ŀ
      SheetsCount := SelSheets.Count;
    end;

    //ÿ�β��Ҷ�Ҫ����ȷ��һ�²��ҷ�Χ
    //�͵������һ���������ǰ������ѡ��Ԫ����ĿΪ1��������������
    //����ֻ�Ƕ�Selection���в���
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end;
    end;

    //������һ�㣬Ĭ���õ�ǰ���ҷ�Χ���в��ң����ϴ����µģ�
    FindRes := Find(FindReplaceRange);

    //�����Ƿ񻻱����
    //����ҲҪ��ô��  ���л�������һ��������
    if FindRes <> nil then begin
      //����ҵ���
      if _FirstFindInSheet then begin
        //�����ڵ�ǰ�������е�һ�η���
        //�ѱ�־��Ϊ�٣������ҵ������ݱ�������
        _FirstFindInSheet := false;
        FirstFindRangeRow := FindRes.Row;
        FirstFindRangeCol := FindRes.Column;
      end
      else begin
        //������ǣ���Ƚ��ҵ��ĺͱ����
        if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
           //���һ�£�˵������һ���ˣ�Ҫ������һ�β��ң��Ѳ��ҽ������
          FindRes := nil;
        end;
      end;
    end;

    if FindRes = nil then begin
       //û�ҵ��������ݵĹ�������Ŀ��ʼ��Ϊ0
      NoFindOutSheetCount := 0;
       //��ȡ��ǰ����������Ϊ����������
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //û�ҵ��������ݹ�������Ŀ�Ѿ�Ϊ������������Ŀ
             //˵��������������û����Ҫ�����ݣ��˳�����
          break;
        end
        else begin
             //�ҵ���һ����������������������1ȡģ��
             //ģΪ������������1����Ϊ������������1��ʼ����ȡģ���Ϊ0��Ҫ����Ϊ1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := SelSheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
                //ѡ��һ���±����µ������ҷ�Χ
                //��ѡ���±�Ȼ�������Χ
                //���˹�����Ҫ����ѡ��״̬������ǰ���Ԫ������һ������
            WorkSheet.Activate;
            WorkSheet.Select(false);

                //�����ǰ��ѡ����������ʲôҲ����
            Selection := FApplication.Selection[$0804];
            if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
              Exit;
            end;

                //�͵������һ���������ǰ������ѡ��Ԫ����ĿΪ1��������������
                //����ֻ�Ƕ�Selection���в���
            if FindReplaceRange.Count = 1 then begin
              FindReplaceRange := WorkSheet.Cells;
            end;

                //��Ӱ����ҵ���ʼλ��
            _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
          end;
        end;

        FindRes := Find(FindReplaceRange);

          //�ж��Ƿ��һ���������һ��
        if FindRes <> nil then begin
            //����ҵ���
          if _FirstFindInSheet then begin
              //�����ڵ�ǰ�������е�һ�η���
              //�ѱ�־��Ϊ�٣������ҵ������ݱ�������
            _FirstFindInSheet := false;
            FirstFindRangeRow := FindRes.Row;
            FirstFindRangeCol := FindRes.Column;
          end
          else begin
              //������ǣ���Ƚ��ҵ��ĺͱ����
            if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
                 //���һ�£�˵������һ���ˣ�Ҫ������һ�β��ң��Ѳ��ҽ������
              FindRes := nil;
            end;
          end;
        end;

        if FindRes = nil then
             //��������ѡȡ��һ��������û���ҵ����򽫱�������Ŀ��1����ʾ�ñ�û�д���������
             //ѭ����Ĳ���ʧ�ܲ���֤�������޲������ݣ��п����ǲ鵽β��
          NoFindOutSheetCount := NoFindOutSheetCount + 1;
        begin
        end
      until (FindRes <> nil);

//       if FindRes <> nil then
//       begin
//          //���ڶ�ѡ������������һ���µĹ���������ȡ����ǰ�������ѡ��
//          WorkSheet.Activate;
//       end;
    end;
  end
  else begin
    //�����ǰֻѡ����һ���������ڵ�һ�β��ҵ�ʱ����Ҫȷ�����ҷ�Χ
    //���ѡ�е���һ����Ԫ����Ҫ��������Ϊ���鷶Χ
    //���ѡ�е����������ϵ�Ԫ�񣬵�ǰ��ѡ��������Ǵ��鷶Χ�����ض��⴦��
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end
    end;

    //ʵ�ʽ��в���
    FindRes := Find(FindReplaceRange);
  end;

  //�Բ��ҵ��Ľ��������
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
  NextSheet: IDispatch; //������ʹ��
  WorkSheet: IWorkSheet;
  FindRes: IETRange;
  RangeTemp: OleVariant;
  NoFindOutSheetCount: Integer; //���ҹ���û���ҵ��������ݵĹ�������Ŀ�����ڹ������Ͷ�ѡ������ʱ�Ĳ���
  SheetIndex: Integer;
  SelSheets: Sheets;
  listitem: TTntListItem;
  NameTemp: IName;
  AddressStr :Widestring;
  AddressTemp :WideString;
  I: Integer; //ѭ������
begin
  // todo: �������ȥ�����棬Ӧ����ϸ����߼���ϵ��
  SheetIndex := 0;
  if FApplication = nil then
    Exit;
  if Assigned(FApplication) then
    SelSheets := FApplication.ActiveWindow.SelectedSheets;

  Selection := FApplication.Selection[$0804];
  if (Selection = nil) then exit;

  //�������������Ϊ�գ���ʾ����
  if mcbFind.Text = '' then begin
    lvFindAll.Clear;
    KSO_MessageBox(WideLoadResString(@sFindReplace_InvalidWhat), '', MB_OK + MB_ICONEXCLAMATION);
    mcbFind.SetFocus;
    Exit;
  end;

  { add by oyxz, see #C1 }
  // �������ƹ����Ĵ��룬���Ϊ��ֵʱ�����������²���������ᵼ��bug:26096
  mcbFind.ActivateCurrItem;
  mcbReplace.ActivateCurrItem;
  { add end }

  //�����ǰ��ѡ����������ʲôҲ����
  if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
    Exit;
  end;

  //ȷ�����ҷ�Χ��ע���е�������ظı�
  //���ڹ�������һ�β���֮�󶼲��ظı�
  if SearchRange = srWorkBook then begin
    //�����ǰ���ҷ�Χ�Ǹ�������
    //����ǵ�һ�β��ң����ȡ��ǰ��������Cells�������ҷ�Χ������Ҫ�����ǰ��Selection
    if not _FindNextFlag then begin
      SheetsCount := FApplication.ActiveWorkbook.Worksheets.Count;
      _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
    end;

    ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
    if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
      FindReplaceRange := WorkSheet.Cells;
       //FApplication.ActiveCell.Select;//ȡ��ԭ�������ϵ�����ѡ��
    end
    else begin
      Exit;
    end;

    //������һ�㣬Ĭ���õ�ǰ���ҷ�Χ���в��ң����ϴ����µģ�
    FindRes := Find(FindReplaceRange);

    //�����Ƿ񻻱����
    //����ҲҪ��ô��  ���л�������һ��������
    if FindRes <> nil then begin
      //����ҵ���
      if _FirstFindInSheet then begin
        //�����ڵ�ǰ�������е�һ�η���
        //�ѱ�־��Ϊ�٣������ҵ������ݱ�������
        _FirstFindInSheet := false;
        FirstFindRangeRow := FindRes.Row;
        FirstFindRangeCol := FindRes.Column;
      end
      else begin
        //������ǣ���Ƚ��ҵ��ĺͱ����
        if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
           //���һ�£�˵������һ���ˣ�Ҫ������һ�β��ң��Ѳ��ҽ������
          FindRes := nil;
        end;
      end;
    end;

    if FindRes = nil then begin
       //û�ҵ��������ݵĹ�������Ŀ��ʼ��Ϊ0
      NoFindOutSheetCount := 0;
       //��ȡ��ǰ����������Ϊ����������
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //û�ҵ��������ݹ�������Ŀ�Ѿ�Ϊ������������Ŀ
             //˵��������������û����Ҫ�����ݣ��˳�����
          break;
        end
        else begin
             //�ҵ���һ����������������������1ȡģ��
             //ģΪ������������1����Ϊ������������1��ʼ����ȡģ���Ϊ0��Ҫ����Ϊ1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := FApplication.ActiveWorkbook.Worksheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
            FindReplaceRange := WorkSheet.Cells;
               //���˹�����Ҫ����ѡ��״̬������ǰ���Ԫ������һ������
               //��Ӱ����ҵ���ʼλ��
            _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
          end;
        end;

        if WorkSheet.Visible <> etSheetVisible then
          FindRes := nil // ����������ɼ����򲻲���
        else
          FindRes := Find(FindReplaceRange);

          //�ж��Ƿ��һ���������һ��
        if FindRes <> nil then begin
            //����ҵ���
            //ֻ�е��ù�������ƥ������ʱ���Ž��ñ�ѡ��
          WorkSheet.Select(true);
            //���������ԭ��
            //1�����һ��ʼ��ѡ�й�������ù�������û�б��������ݾͻ������˸���ñ�һ��������
            //2�������ʼ��ѡ�й�������ң���ǰ��Ԫ�񽫻����ϸ�������ģ������±����λ��������
            //���Բ��õķ��������ң�����ҵ�����ѡ������
          Find(FindReplaceRange);
          if _FirstFindInSheet then begin
              //�����ڵ�ǰ�������е�һ�η���
              //�ѱ�־��Ϊ�٣������ҵ������ݱ�������
            _FirstFindInSheet := false;
            FirstFindRangeRow := FindRes.Row;
            FirstFindRangeCol := FindRes.Column;
          end
          else begin
              //������ǣ���Ƚ��ҵ��ĺͱ����
            if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
                 //���һ�£�˵������һ���ˣ�Ҫ������һ�β��ң��Ѳ��ҽ������
              FindRes := nil;
            end;
          end;
        end;

        if FindRes = nil then
             //��������ѡȡ��һ��������û���ҵ����򽫱�������Ŀ��1����ʾ�ñ�û�д���������
             //ѭ����Ĳ���ʧ�ܲ���֤�������޲������ݣ��п����ǲ鵽β��
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
    //�����ǰѡ���˶��������
    if not _FindNextFlag then begin
      _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
      //�˴���¼�Ĺ�������Ŀ�ǵ�ǰѡ�еĹ�������Ŀ
      SheetsCount := SelSheets.Count;
    end;

    //ÿ�β��Ҷ�Ҫ����ȷ��һ�²��ҷ�Χ
    //�͵������һ���������ǰ������ѡ��Ԫ����ĿΪ1��������������
    //����ֻ�Ƕ�Selection���в���
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end;
    end;

    //������һ�㣬Ĭ���õ�ǰ���ҷ�Χ���в��ң����ϴ����µģ�
    FindRes := Find(FindReplaceRange);

    //�����Ƿ񻻱����
    //����ҲҪ��ô��  ���л�������һ��������
    if FindRes <> nil then begin
      //����ҵ���
      if _FirstFindInSheet then begin
        //�����ڵ�ǰ�������е�һ�η���
        //�ѱ�־��Ϊ�٣������ҵ������ݱ�������
        _FirstFindInSheet := false;
        FirstFindRangeRow := FindRes.Row;
        FirstFindRangeCol := FindRes.Column;
      end
      else begin
        //������ǣ���Ƚ��ҵ��ĺͱ����
        if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
           //���һ�£�˵������һ���ˣ�Ҫ������һ�β��ң��Ѳ��ҽ������
          FindRes := nil;
        end;
      end;
    end;

    if FindRes = nil then begin
       //û�ҵ��������ݵĹ�������Ŀ��ʼ��Ϊ0
      NoFindOutSheetCount := 0;
       //��ȡ��ǰ����������Ϊ����������
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //û�ҵ��������ݹ�������Ŀ�Ѿ�Ϊ������������Ŀ
             //˵��������������û����Ҫ�����ݣ��˳�����
          break;
        end
        else begin
             //�ҵ���һ����������������������1ȡģ��
             //ģΪ������������1����Ϊ������������1��ʼ����ȡģ���Ϊ0��Ҫ����Ϊ1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := SelSheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
                //ѡ��һ���±����µ������ҷ�Χ
                //��ѡ���±�Ȼ�������Χ
                //���˹�����Ҫ����ѡ��״̬������ǰ���Ԫ������һ������
            WorkSheet.Activate;
            WorkSheet.Select(false);

                //�����ǰ��ѡ����������ʲôҲ����
            Selection := FApplication.Selection[$0804];
            if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
              Exit;
            end;

                //�͵������һ���������ǰ������ѡ��Ԫ����ĿΪ1��������������
                //����ֻ�Ƕ�Selection���в���
            if FindReplaceRange.Count = 1 then begin
              FindReplaceRange := WorkSheet.Cells;
            end;

                //��Ӱ����ҵ���ʼλ��
            _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
          end;
        end;

        FindRes := Find(FindReplaceRange);

          //�ж��Ƿ��һ���������һ��
        if FindRes <> nil then begin
            //����ҵ���
          if _FirstFindInSheet then begin
              //�����ڵ�ǰ�������е�һ�η���
              //�ѱ�־��Ϊ�٣������ҵ������ݱ�������
            _FirstFindInSheet := false;
            FirstFindRangeRow := FindRes.Row;
            FirstFindRangeCol := FindRes.Column;
          end
          else begin
              //������ǣ���Ƚ��ҵ��ĺͱ����
            if (FindRes.Row = FirstFindRangeRow) and (FindRes.Column = FirstFindRangeCol) then begin
                 //���һ�£�˵������һ���ˣ�Ҫ������һ�β��ң��Ѳ��ҽ������
              FindRes := nil;
            end;
          end;
        end;

        if FindRes = nil then
             //��������ѡȡ��һ��������û���ҵ����򽫱�������Ŀ��1����ʾ�ñ�û�д���������
             //ѭ����Ĳ���ʧ�ܲ���֤�������޲������ݣ��п����ǲ鵽β��
          NoFindOutSheetCount := NoFindOutSheetCount + 1;
        begin
        end
      until (FindRes <> nil);

//       if FindRes <> nil then
//       begin
//          //���ڶ�ѡ������������һ���µĹ���������ȡ����ǰ�������ѡ��
//          WorkSheet.Activate;
//       end;
    end;
  end
  else begin
    //�����ǰֻѡ����һ���������ڵ�һ�β��ҵ�ʱ����Ҫȷ�����ҷ�Χ
    //���ѡ�е���һ����Ԫ����Ҫ��������Ϊ���鷶Χ
    //���ѡ�е����������ϵ�Ԫ�񣬵�ǰ��ѡ��������Ǵ��鷶Χ�����ض��⴦��
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end
    end;

    //ʵ�ʽ��в���
    FindRes := Find(FindReplaceRange);
  end;

  //�Բ��ҵ��Ľ��������
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
    lvFindAll.Clear;//��ձ�
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
  NoFindOutSheetCount: Integer; //���ҹ���û���ҵ��������ݵĹ�������Ŀ�����ڹ������Ͷ�ѡ������ʱ�Ĳ���
  SheetIndex: Integer;
  NextSheet: IDispatch; //������ʹ��
  fIsReplaceFail: Boolean;
  SelSheets: Sheets;
begin
  SheetIndex := 1; //���������岢����ֻ��Ϊ���ٳ���һ����might not have been initialized����Warning
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

    //�������������Ϊ�գ���ʾ����
    if mcbFind.Text = '' then begin
      KSO_MessageBox(WideLoadResString(@sFindReplace_InvalidWhat), '', MB_OK + MB_ICONEXCLAMATION);
      mcbFind.SetFocus;
      Exit;
    end;
  end;

  //�����ǰ��ѡ����������ʲôҲ����
  if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
    Exit;
  end;

  //ȷ�����ҷ�Χ��ע���е�������ظı�
  //���ڹ�������һ�β���֮�󶼲��ظı�
  //
  if SearchRange = srWorkBook then begin
    //�����ǰ���ҷ�Χ�Ǹ�������
    //����ǵ�һ�β��ң����ȡ��ǰ��������Cells�������ҷ�Χ������Ҫ�����ǰ��Selection
    if _ReplaceFirstFlag then begin
      SheetsCount := FApplication.ActiveWorkbook.Worksheets.Count;
      _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
    end;

    ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
    if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
      FindReplaceRange := WorkSheet.Cells;
      FApplication.ActiveCell.Select; //ȡ��ԭ�������ϵ�����ѡ��
    end
    else begin
      Exit;
    end;

    if WorkSheet.Visible <> etSheetVisible then
      FindRes := nil  // ������������ز��滻
    else
      //������һ�㣬Ĭ���õ�ǰ���ҷ�Χ���в��ң����ϴ����µģ�
      FindRes := Replace(FindReplaceRange, fIsReplaceFail);

    if FindRes = nil then begin
       //����ڵ�ǰ������û���ҵ�����Ҫ���ǵ���һ�������������
       //û�ҵ��������ݵĹ�������Ŀ��ʼ��Ϊ0
      NoFindOutSheetCount := 0;
       //��ȡ��ǰ����������
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //�ڹ�������ѭ������
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //û�ҵ��������ݹ�������Ŀ�Ѿ�Ϊ������������Ŀ
             //˵��������������û����Ҫ�����ݣ��˳�����
          break;
        end
        else begin
             //�ҵ���һ����������������������1ȡģ��
             //ģΪ������������1����Ϊ������������1��ʼ����ȡģ���Ϊ0��Ҫ����Ϊ1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := FApplication.ActiveWorkbook.Worksheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
            FindReplaceRange := WorkSheet.Cells;
               //���˹�����Ҫ����ѡ��״̬������ǰ���Ԫ������һ������
               //��Ӱ����ҵ���ʼλ��
            _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
          end;
        end;

        if WorkSheet.Visible <> etSheetVisible then
          FindRes := nil  // ������������ز��滻
        else
          FindRes := Find(FindReplaceRange);

        if FindRes = nil then begin
             //��������ѡȡ��һ��������û���ҵ����򽫱�������Ŀ��1����ʾ�ñ�û�д���������
             //ѭ��ǰ�Ĳ���ʧ�ܲ���֤�������޲������ݣ��п��ܵ�һ��������鵽β��
          NoFindOutSheetCount := NoFindOutSheetCount + 1;
        end
        else begin
             //ֻ�е��ù�������ƥ������ʱ���Ž��ñ�ѡ��
          WorkSheet.Select(true);
             //����һ�飬ԭ��͹���������ͬ��
          FindRes := Find(FindReplaceRange);
        end
      until (FindRes <> nil);
    end;
  end
////////////////////////////////////////////////////////////////////////////
  else if SelSheets.Count > 1 then begin
    //�����ǰѡ���˶��������
    if _ReplaceFirstFlag then begin
      SheetsCount := SelSheets.Count;
      _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
    end;

    //ÿ�β��Ҷ�Ҫ����ȷ��һ�²��ҷ�Χ
    //�͵������һ���������ǰ������ѡ��Ԫ����ĿΪ1��������������
    //����ֻ�Ƕ�Selection���в���
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end;
    end;

    //������һ�㣬Ĭ���õ�ǰ���ҷ�Χ���в��ң����ϴ����µģ�
    FindRes := Replace(FindReplaceRange, fIsReplaceFail);

    if FindRes = nil then begin
       //����ڵ�ǰ������û���ҵ�����Ҫ���ǵ���һ�������������
       //û�ҵ��������ݵĹ�������Ŀ��ʼ��Ϊ0
      NoFindOutSheetCount := 0;
       //��ȡ��ǰ����������
      ActiveSheet := FApplication.ActiveSheet; //ActiveSheet�Ǹ�Dispatch
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        SheetIndex := WorkSheet.Index[$0804];
      end;

       //�ڹ�������ѭ������
      repeat
        if NoFindOutSheetCount = SheetsCount then begin
             //û�ҵ��������ݹ�������Ŀ�Ѿ�Ϊ������������Ŀ
             //˵��������������û����Ҫ�����ݣ��˳�����
          break;
        end
        else begin
             //�ҵ���һ����������������������1ȡģ��
             //ģΪ������������1����Ϊ������������1��ʼ����ȡģ���Ϊ0��Ҫ����Ϊ1
          SheetIndex := (SheetIndex + 1) mod (SheetsCount + 1);
          if SheetIndex = 0 then begin
            SheetIndex := 1;
          end;

          NextSheet := SelSheets[SheetIndex];
          if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
                //ѡ��һ���±����µ������ҷ�Χ
                //��ѡ���±�Ȼ�������Χ
                //���˹�����Ҫ����ѡ��״̬������ǰ���Ԫ������һ������
            WorkSheet.Activate;
            WorkSheet.Select(false);

                //�����ǰ��ѡ����������ʲôҲ����
            Selection := FApplication.Selection[$0804];
            if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
              Exit;
            end;

                //�͵������һ���������ǰ������ѡ��Ԫ����ĿΪ1��������������
                //����ֻ�Ƕ�Selection���в���
            if FindReplaceRange.Count = 1 then begin
              FindReplaceRange := WorkSheet.Cells;
            end;

                //��Ӱ����ҵ���ʼλ��
            _FirstFindInSheet := true; //ÿ�θĻ�sheet��Ҫ�Ѹñ�־��ΪTRUE
          end;
        end;

        FindRes := Find(FindReplaceRange);

        if FindRes = nil then
             //��������ѡȡ��һ��������û���ҵ����򽫱�������Ŀ��1����ʾ�ñ�û�д���������
             //ѭ��ǰ�Ĳ���ʧ�ܲ���֤�������޲������ݣ��п��ܵ�һ��������鵽β��
          NoFindOutSheetCount := NoFindOutSheetCount + 1;
        begin
        end
      until (FindRes <> nil);
    end;
  end
  else begin
    //�����ǰֻѡ����һ��������
    //���ѡ�е���һ����Ԫ����Ҫ��������Ϊ���鷶Χ
    //���ѡ�е����������ϵ�Ԫ�񣬵�ǰ��ѡ��������Ǵ��鷶Χ�����ض��⴦��
    if IsSingleCell(FindReplaceRange) then begin
      ActiveSheet := FApplication.ActiveSheet;
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end
    end;

    //ʵ�ʽ����滻
    FindRes := Replace(FindReplaceRange, fIsReplaceFail);
  end;

  if FindRes <> nil then
    FindRes.Activate
  else begin
    if FbProtected then
      // ����ʱ�������ں��Ѿ�������
    else if fIsReplaceFail = true then
      //�滻�����з�������
      KSO_MessageBox(WideLoadResString(@sReplace_ReplaceFailed), '', MB_OK + MB_ICONEXCLAMATION)
    else
      //�Ҳ���ƥ����
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
  ReplaceAllCount: Integer; //�ܵ�ȫ���滻����Ŀ
  SheetReplaceAllCount: Integer; //������ȫ���滻����Ŀ
  SheetsCount: Integer; //�������ڹ��������Ŀ
  I: Integer; //ѭ������
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

  //�������������Ϊ�գ���ʾ����
  if mcbFind.Text = '' then begin
    KSO_MessageBox(WideLoadResString(@sFindReplace_InvalidWhat), '', MB_OK + MB_ICONEXCLAMATION);
    mcbFind.SetFocus;
    Exit;
  end;

  //�����ǰ��ѡ����������ʲôҲ����
  if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
    Exit;
  end;

  //ȷ�����ҷ�Χ
  if SearchRange = srWorkBook then begin
    //�����ǰ���ҷ�Χ�Ǹ�������
    //�����������ȫ���滻����¼�滻������

    //ȷ�����������м�����������ʼ�������滻���
    SheetsCount := AllSheets.Count;
    for I := 1 to SheetsCount do begin
      //��ȡ������
      NextSheet := AllSheets[I];
      if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        if WorkSheet.Visible <> etSheetVisible then continue;
		if WorkSheet.ProtectContents then
               continue;
          //��������������Ϊ���ҷ�Χ
        FindReplaceRange := WorkSheet.Cells;
          //�Ըù��������ȫ���滻
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
    //�����ǰѡ���˶��������
    SheetsCount := SelSheets.Count;
    for I := 1 to SheetsCount do begin
      //��ȡ������
      NextSheet := SelSheets[I];
      if NextSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
		if WorkSheet.ProtectContents then
                continue;
        //����
        WorkSheet.Activate;
        WorkSheet.Select(false);

        //��ȡselection����֣����������������һ�ο����õ����ڶ�������������Ŀ��Ϊ0�ˣ�
        Selection := FApplication.Selection[$0804];
        if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
          Exit;
        end;

        if IsSingleCell(FindReplaceRange) then begin
           //��������������Ϊ���ҷ�Χ
          FindReplaceRange := WorkSheet.Cells;
        end
        else begin
           //�����ǰ��ѡ������
          Selection := FApplication.Selection[$0804];
          if Selection.QueryInterface(IETRange, FindReplaceRange) <> S_OK then begin
            Exit;
          end;
        end;

        //�Ըù��������ȫ���滻
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
    //�����ǰֻѡ����һ��������
    if IsSingleCell(FindReplaceRange) then begin
      //���ѡ�е���һ����Ԫ����Ҫ��������Ϊ���鷶Χ
      //���ѡ�е����������ϵ�Ԫ�񣬵�ǰ��ѡ��������Ǵ��鷶Χ�����ض��⴦��
      ActiveSheet := FApplication.ActiveSheet;
      if ActiveSheet.QueryInterface(IWorkSheet, WorkSheet) = S_OK then begin
        FindReplaceRange := WorkSheet.Cells;
      end
    end;

    //����ȫ���滻
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
      //�滻�����з�������
    if ReplaceAllCount <> -2 then  // ���ں�Լ����ʾʲô��û��
      KSO_MessageBox(WideLoadResString(@sReplace_ReplaceFailed), '', MB_OK + MB_ICONEXCLAMATION);
    btnReplaceAll.SetFocus;
  end;
end;

procedure TFindReplaceDlg.OnEnterSearch(Sender: TObject);
begin
  // ET�鷴ӳ��͸��Ӱ�����ܣ���ע�͡�2004-7-11 -tifi
//  AlphaBlend := True;
//  AlphaBlendValue := 200;
end;

procedure TFindReplaceDlg.OnExitSearch(Sender: TObject);
begin
  // ET�鷴ӳ��͸��Ӱ�����ܣ���ע�͡�2004-7-11 -tifi
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
  //����
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
   //���ڵ����滻
   //�����滻��ǰ��Ԫ��Ȼ���ҵ���һ�����滻�ĵ�Ԫ������Ϊ��ǰ��Ԫ��
   //�����Ļ��������ǰ��Ԫ�����滻�ɹ�����һ�¡��滻���Ľ���ǳ����ҵ���һ�����滻����
  ReplaceCount := FApplication.ActiveCell.Replace(WideString(mcbFind.Text), WideString(mcbReplace.Text), GetFindLookAt,
    GetSearchOrder, GetMatchCase, GetMatchByte, NULL, NULL);

  if ReplaceCount > 0 then begin
      //�����ǰ��Ԫ���滻�ɹ�
      //����һ��
    if _ReplaceFirstFlag then begin
       //����ǵ�һ���滻�����ں�������Find
      FindRes := FRange.Find(WideString(mcbFind.Text), FApplication.ActiveCell, GetFindLookin,
        GetFindLookAt, GetSearchOrder, etNext, GetMatchCase, GetMatchByte, NULL);
      _ReplaceFirstFlag := false;
    end
    else begin
       //������FindNext
      FindRes := FRange.FindNext(FApplication.ActiveCell);
    end;

      //ȷ������ʲô
    if FindRes <> nil then begin
        //�����ҵ���һ�����򷵻���һ��
      Result := FindRes;
    end
    else begin
        //��֮���ص�ǰ��Ԫ��
      Result := FApplication.ActiveCell;
    end
  end
  else if ReplaceCount = 0 then begin
      //�����ǰ��Ԫ��ƥ�䣬���û���滻����
      //ֱ������һ������
    if _ReplaceFirstFlag then begin
       //����ǵ�һ���滻�����ں�������Find
      Result := FRange.Find(WideString(mcbFind.Text), FApplication.ActiveCell, GetFindLookin,
        GetFindLookAt, GetSearchOrder, etNext, GetMatchCase, GetMatchByte, NULL);
      _ReplaceFirstFlag := false;
    end
    else begin
       //������FindNext
      Result := FRange.FindNext(FApplication.ActiveCell);
    end;
  end
  else if ReplaceCount = -2 then begin
    // ������滻ʧ�ܣ���Ҫ��������ݿ��ܲ��ǺϷ��Ĺ�ʽ�����߳��ڱ���״̬
    // -2�Ǻ�CoreԼ����ֵ
    FbProtected := True;
    Result := nil;
    end
  else if ReplaceCount = -1 then begin
    Result := nil;
    fIsReplaceFail := True;
    end
  else begin
    // �������������
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
  FirstItem: IDispatch; //�ڲ�������count��Ϊ1ʱʹ��
  FirstCell: IETRange;
  MergeArea: IETRange;
  MergeAreaAddress: WideString; //�����ϲ��жϵ�����ĵ�ַ
  RangeAddress: WideString; //ԭʼ��������ĵ�ַ
  RelativeTo: OleVariant;
begin
  //������������countΪ1����϶��ǵ�����Ԫ��
  //���������count��Ϊ1����Ҫ�ж��Ƿ�����Ϊһ���ϲ���Ԫ��
  //�����ǣ�
  //�Ӹ�������ȡһ����Ԫ�񡣿���ʹ��item(1,1)
  //�Ըõ�Ԫ���� MergeArea ��Ȼ�����������Ƿ��ԭ����һ�£����һ�£���˵�����������Ǹ��ϲ���Ԫ��
  Result := false;

  if FRange.Count = 1 then begin
    Result := true;
  end
  else begin
    //���������count��Ϊ1
    FirstItem := FRange.item[1, 1];
    if FirstItem.QueryInterface(IETRange, FirstCell) = S_OK then begin
      MergeArea := FirstCell.MergeArea;
      //������������ñȽ����⣬�Ժ�Ҫ�ģ�Ŀǰ�⼸������APIʵ����Ҳû�д���
      MergeAreaAddress := MergeArea.Address[false, false, etA1, false, RelativeTo];
      RangeAddress := FRange.Address[false, false, etA1, false, RelativeTo];

      if MergeAreaAddress = RangeAddress then begin
        //����ϲ���Ԫ��ĵ�ַ���ڲ�����ַ
        //��˵�������Ǹ��ϲ���Ԫ��
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
      // ��ʾ���ɶ�λ��Ϣ
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
  // ��ʱ���ε������ӹ���
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
  ActiveWorkbook: IDispatch; //������ʹ��
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
  //��ʱ������
  ColumnToSort := Column.Index;
  lvFindAll.AlphaSort;
  asc := not asc;
end;

procedure TFindReplaceDlg.OnClick(Sender: TObject);
begin
  //��ʱ������
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
