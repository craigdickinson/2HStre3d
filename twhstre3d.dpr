{ **************************************************************************** }
{  Name    : twhstre3d.dpr
'
'  Author  : Richard Harrison
'  Notice  : Copyright (c) 2009 2H Offshore Engineering Ltd
'          : All Rights Reserved
'
'  Notes   : Stress check from Flexcom
'
'
' CVS VERSION INFO
'
'      Last Checked in by: $Author: harrisor $
'      Last Checked in:    $Date: 2015/09/16 17:30:09 $
'              Filename:   $RCSfile: twhstre3d.dpr,v $
'              Revision:   $Revision: 1.96 $
'              Tag:        $Name:$                                             }
{ **************************************************************************** }

program twhstre3d;

{$APPTYPE CONSOLE}

{ Written by 2H OFFSHORE ENGINEERING LIMITED

  This program reads data from the Flexcom .tab and .out files and
  postprocesses to produce a table of 3D results

  REVISION HISTORY

  Version   Author        Date      Description
  ------------------------------------------------------------------------------
  1.0       N.Alderton    07.02.95
  1.1       N.Alderton    01.03.95  Pressure effects removed from effective tension
  1.2       N.Alderton    21.03.95  Buoyancy diameter used in hoopstress calculation
  1.3       N.Alderton    22.03.95  Nonlinear spring stroke corrected
  1.4       N.Alderton    04.04.95  Multiple nonlinear springs incorporated
  1.5       N.Alderton    03.05.95  Unlimited elements, factors read from .INF file
  1.6       N.Alderton    10.05.95  Procedure nodeheight optimised for improvement in speed
  1.7       N.Alderton    16.05.95  Numbers separated by a space in displacement table
  1.8       N.Alderton    20.06.95  Articulation table added
  1.9       N.Alderton    21.06.95  Zero stiffness articulations and zero intsernal diameter added,
  1.10      N.Alderton    07.08.95  Buoyancy diameter changed to de in in hoopstress and effective tension calculation
  1.11      N.Alderton    12.09.95  Axial stress due to internal pressure duplication removed
  1.12      N.Alderton    13.09.95  Thick wall hoopstress theory added
  1.13      N.Alderton    04.10.95  Axial stress due to internal pressure removed completely to correct 1.11
  1.14      N.Alderton    27.10.95  Stress due to torque corrected
  1.15      N.Alderton    19.03.96  Actual OD used, not equivalenced OD
  1.16      N.Alderton    22.07.96  Revised to accept different output for Flexcom ver 3.2.3
  1.17      N.Alderton    06.02.97  Bend radius columns added
  1.18      N.Alderton    03.10.97  Bend radius calcs fixed to take 0 curvature
  1.19      T.Eyles       16.02.98  Tension correction problem fixed (not reading tension correction)
  1.20      M.Campbell    17.12.98  Option for API-RP-2RD mid-wall stress calculation
  1.21      M.Campbell    01.04.99  Updated to allow old INF files to still be used and no. of springs increased to 100
  1.22      M.Campbell    04.05.99  Multiple fluids incorporated, tension correction updated
  1.23      M.Campbell    05.10.99  Version 5 spring reading debugged
  1.24      D.Thomas      28.09.01  Output for bending, axial, hoop and radial stress added (*.TB2 file)
  1.25      C.Bridge      16.01.02  Modified code to allow models with over 1000 elements (added FirstWord procedure)
  1.25.1    C.Bridge      16.01.02  Minor update to account for spring properties and fix the API space bug
  1.26      M.Campbell    17.01.02  Modified to include time domain stress checking option
  1.27      M.Campbell    12.02.02  Revised to accept Flexcom ver 6.1.1 output format
  1.27.1    M.Campbell    01.05.02  Floating point error fix, curvature made double instead of real
  1.29      M.Podskarbi   23.07.02  Code updated to work under Delphi, nonlinear springs stroke bug fixed, tb2 format changed, stt bug fixed
  1.30      M.Podskarbi   21.10.02  Bug in CompileTB2 changed, to allow to process more than 999 elements
  1.31      C.Bridge      22.10.02  Minor update to errorCk Internal Fluid Data Sets - and add Local Varibles for FOR loops (Delphi gives a warning amd does not compile)
                                    Additionally a link to STRUnit has been added - FirstWord routine removed
  1.32      D.Lee         20.01.03  Extract articulation angles using *.tab file, two decimal places for angle output, restore GotoXY interface
                                    and formatting spring forces, component stress calculation with internal and external corrosion corrected
  1.33      D.MacLure     26.03.03  Added code in the Quant and Geometry procedures so that it doesn't crash when bending stiffeners are present in the output file.
  1.34      D.Maclure     09.06.03  Changed the spring from KN to N
  1.35      C.Griffin     11.07.03  Fixed formatting error with spring/articulation section; articulations are now on their own page
                                    Nonlinear elements will no longer crash the geometry procedure.
  1.40      M.Campbell    03.10.03  Removed Version 1.33 bend stiffener additions and changed a few lines of code to work for bend stiffeners (lines 1119 and 1206)
  1.41      M.Campbell    16.06.04  Modified *.sp1 output to allow silly element numbering system that is incompatible with FLEXCOM TAB file output (lines 293 and 302)
  1.42      C.Bridge      13.10.04  Updated tb2 file so bending and axial stresses output correspond to the maximum vM
  1.43      C.Bridge      31.01.05  Updated for FLEXCOM version 7 - Updates to: springs, articulations, and node reading from TAB files (different length headers)
  1.44      C.Bridge      13.05.05  Fixed error from articulations with 0.0 stiffness.
  1.45      C.Bridge      26.10.05  Added 'divide by zero' check in 'springs' routine
  1.46      C.Bridge      17.11.05  Updated to handle unlimited filename lengths
  1.47      M.Campbell    26.01.06  Fixed Internal pressure error.
            & C.Bridge
  1.48      R.I.Harrison  10.07.06  Updates for inclusion of DNV combined stress checks:
                                    - additionally fixed error in springs rountine read only data from current spring
                                    - min and max error in bending radius fixed
  1.49      R.I.Harrison  26.03.07  Updated to allow non-linear spring read in from Modificiation version 7.3 of Flexcomm
  1.49a     R.I.Harrison  27.08.08  Updated to allow version not in capital for Flexcom 7.7.2
  1.50      A.Ghatak      10.03.09  1. Changed formulation for application of tension factor and tension correction to global tension.
                                    2. Introduced optional keyword "NOMINAL STATIC TENSION" to fit into the new formulation.
                                    3. Introduced optional keywords "NOMINAL ID", "INTERNAL FLUID" and "EXTERNAL FLUID" to allow user to input ids and fluid properties.
                                    4. Modified procedure 'Quan' to differentiate between nonlinear springs and nonlinear material properties.
                                    5. Modified procedure 'dowrite' to read in D, G and I values differently.
                                       This is to make sure if the number of values for G or I is entered incorrectly, 2HSTRE3D gives a warning to prevent incorrect calculations.
                                    6. Modified procedure 'sttloads' to include tension correction calculations.
                                    7. Procedure 'vmtimetrace' added to output the vM stress timetrace for a requested element. New optional keyword 'VMTT' added in procedure 'dowrite'.
                                    8. A new optional keyword 'ZERO PRESSURE CHECK' is introduced. It gives the option to check which vM stress is higher - with top pressure or without top pressure.
                                       It outputs the higher vM stress value if zero pressure check is activated.
  1.51      B.Yue         15.05.09  Added the ISO creteria for completion/work over risers
  1.51.1.0  R.Harrison    10.08.09  Added syntax format in INF to catch an syntax error
  1.51.2.2  R.Harrison    20.10.10  Fixed STT Bug on the error due to stress reporting for SPRING and ARTICS
  1.51.2.3  R.Harrison    01.02.11  Fixed non-linear spring storke results for version Flexcom 7.9 and above
  1.51.2.0  R.Harrison    11.04.11  Release of latest version.
  1.51.3.1  C.Dickinson   04.05.12  - WIP upgrade for Flexcom v8: Added a number of standard v7 -> v8 table format changes.
                                    - Modified GetModelDetails (previously Quan) to read nonlinear flex joint data from .out file.
  1.51.3.2  C.Dickinson   11.05.12  - Rewritten and corrected artics procedure (renamed to FlexJoints) for linear flex joints and added
                                      nonlinear flex joints functionality.
                                    - LinInterp procedure added for performing linear interpolation in processing nonlinear flex joints.
                                    - Updated GetModelDetails including making linear and nonlinear flex joints arrays dynamic and
                                      combining into a single flex joints array.
  1.51.3.3  C.Dickinson   16.05.12  - Modified SSTLoads to handle the change in format of .grd files in Flexcom v8.
                                    - Removed the use of the .art temporary file for articulation/flex joints. All flex joint data is
                                      now stored in an array and written directly to the .tb1 file.
                                    - Added collection of warnings feature. A number of warnings are collected in an array and output
                                      at the end of .tb1 file.
                                    - Warning defintions have also been grouped together in the beginning of the code for convenience.
                                    - General code maintenance to provide consistent clarity of structure and variable names.
  1.51.3.4  C.Dickinson   21.09.12  - Update for VMTT routines to allow correct reading of .grd files for Flexcom v8.
                                    - Fixed issue when handling models which have only nonlinear flex joints and no linear ones.
                                    - More efficient re-write of CompileTB2 plus general minor bug fixes and code tidying.
  1.51.3.0  R.Harrison    18.10.12  - Release.
  1.51.4.1  R.Harrison    22.04.13  - Correct 4 digits limit on the tb2 file output.
  1.51.4.2  R.Harrison    28.05.13  - Fixed Array limitation in VMTT option
  1.51.4.3  R.Harrison    29.05.13  - Fixed for association problem with 100 or more nonlinear springs
  1.51.4.4  R.Harrison    10.06.13  - Add a Coating OD specification in INF to account for correct axial stress under buoyancy
  1.51.4.5  R.Harrison    19.06.13  - Data Echo of Coating OD and warning negative effective in DNV, ISO
  1.51.4.6  R.Harrison    03.07.13  - Fix UpperCase converting of Key filename
  1.51.4.7  R.Harrison    09.07.13  - Add error message on inconsistency GRD timetraces
  1.51.4.0  R.Harrison    06.08.13  - Release of version 1.51.4.0
  1.51.5.1  R.Harrison    27.08.13  - Add catch to Fluid Numbers
  1.51.5.2  R.Harrison    08.10.13  - Fixed Bracket location in BSI formulation
  1.52.0.1  T.Yang        11.10.13  - Add API-STR-2RD
  1.52.0.0  T.Yang        13.02.14  - Release
  1.52.1.1  T.Yang        20.02.14  - Change input format according to new requirement
  1.52.1.0  T.Yang        28.02.14  - Release
  1.52.2.0  T.Yang        03.03.14  - Fix memory allocate issue
  1.52.3.0  T.Yang        11.03.14  - Fix bug of STT option
  1.52.3.1  T.Yang        23.06.14  - Fix bug for bending strain calculation
  1.52.3.2  T.Yang        03.07.14  - Not output bending strain in .tb4 file and output scientific bending radius into .tb5
  1.52.3.3  T.Yang        14.07.14  - Add bending strain based on new request
  1.52.3.4  R.Harrison    27.08.14  - DNV Mod, Corrrosion to OD and strain hardening at actual pressures
  2.0.0.1   R.Harrison    28.08.14  - Remove Temp files
  2.0.0.2   R.Harrison    04.09.14  - Add in the Labels
  2.0.0.3   R.Harrison    08.09.14  - Create Excel Output
  2.0.0.4   R.Harrison    17.09.14  - Label Data in Excel Output
  2.0.0.5   R.Harrison    08.10.14  - Sort out API-STR-2RD echo and stress output to Excel
  2.0.0.6   R.Harrison    13.11.14  - Sort out Time Domain stress by merging STTLoads and VMTimetrace
  2.0.0.7   R.Harrison    11.12.14  - Final tweaks
  2.0.0.8   R.Harrison    05.01.15  - Additional Angualar Label checks
  2.0.0.9   R.Harrison    13.02.15  - Additional Error check in Tab file data entry
  2.0.0.10  R.Harrison    23.02.15  - Data Entry checks + DNV Alternative + TB5 in stress breakdown
  2.0.0.11  R.Harrison    20.04.15  - Update for API STD Table Sheet Delete
  2.0.0.12  R.Harrison    23.04.15  - Correct Label Index Error
  2.0.0.13  R.Harrison    30.04.15  - Add in pressure warning, fix typos
  2.0.0.14  R.Harrison    06.05.15  - Fix Temperature keyword, STT and GRD file not found
  2.0.0.15  R.Harrison    08.05.15  - Check for GRD file, remove old STT element verification
  2.0.0.16  R.Harrison    11.05.15  - Fix array for nonlinear list and timetrace array
  2.0.0.17  R.Harrison    14.05.15  - TensileStrYr applied to the derating
  2.0.0.18  R.Harrison    20.05.15  - Initial Ovality extract in APISTD2RDCalc
  2.0.0.19  R.Harrison    20.05.15  - Pressure to internal pressure in APISTDCheck
  2.0.0.20  R.Harrison    08.06.15  - External/Internal Pressure Override Bug Fix
  2.0.0.21  R.Harrison    03.07.15  - Make BSIShearInput non case sensitive for Time Domain
  2.0.0.22  R.Harrison    12.08.15  - Internal Fluid entry correction + Shear force max to abs for BSI
  2.0.0.23  R.Harrison    13.08.15  - BSI Od - correction to diamout in STT
  2.0.0.24  R.Harrison    24.08.15  - Increase array size for STD with Labels
  2.0.0.25  R.Harrison    08.09.15  - BSI Time domian shear stress extra sqr on shear load
  2.0.0.26  R.Harrison    11.09.15  - No output of method4 data if curvature no read via it
  2.0.0.27  R.Harrison    16.09.15  - Add data echo of function and environmental file to Excel
  2.0.0.28  R.Harrison    18.09.15  - Fix DNV code check error
  2.0.0.29  R.Harrison    28.09.15  - More decimals on Curvature Raduis Data Echo
  2.0.0.30  R.Harrison    13.11.15  - Tb3 header + Ovality
  2.0.0.31  R.Harrison    15.02.16  - Release of version 2.0.0.0
  2.0.1.1   R.Harrison    26.02.16  - Extrapolation option for Flexjoint
  2.0.1.0   R.Harrison    22.03.16  - Release of version 2.0.1.0
  2.0.2.1   R.Harrison    10.06.16  - Modifcation for Flexcom 8.4.4
  2.0.2.2   R.Harrison    12.07.16  - Additional fixes for Flexcom 8.4.4
  2.0.2.0   R.Harrison    13.07.16  - Release
  2.0.3.1   R.Harrison    27.01.17  - Include INF file in output file name
  2.0.3.0   R.Harrison    07.03.17  - Release
  2.0.4.1   R.Harrison    14.03.17  - Allow extremal data input from Flexcom
  2.0.4.2   R.Harrison    07.08.17  - Allow more label data output
  2.0.4.3   R.Harrison    15.09.17  - Label not found to be a warning rather error
  2.0.4.4   R.Harrison    10.10.17  - Update nonlinear spring to deal with non monotonic data
  2.0.4.5   C.Dickinson   17.10.17  - Re-structure of Labels tab
                                                                               }
{ **************************************************************************** }

uses
  Windows,
  SysUtils,
  StrUtils,
  StrUnit,
  Console,
  Math,
  Dialogs,
  ComObj,
  ShellAPI,
  ActiveX,
  Variants;

const
  PROG_NAME = '2HSTRE3D';
  VERSION = '2.0.4.5: 17th October 2017';
  NOT_FOUND = 'Not found';
  MAX_FLUIDS = 20;
  MAX_ELEMENTS = 20000;
  MAX_TIMETRACES = 80000;
  MAX_NODES = 20000;
  MAX_SPRINGS = 10000;
  MAX_LABELS = 6000;
  MAX_MATERIAL_POINTS  = 200;
  GRAVITY = 9.81;
  pi = 3.1415926535;

  //----------
  // Warnings
  //----------
  WARN_NO_ADACENT_ELEMENT = 'Flex joint data not available due to no adjacent element';
  WARN_NO_ANGLE_TIMETRACES = 'Note: No Angle Timetraces in .tab file - check .pst file';
  WARN_ZERO_STIFFNESS = '  Flex joint data not available due to zero stiffness';
  WARN_LIN_INTERP = ' Warning: Linear interpolation failed for nonlinear flex joint calculation - Using Extrapolation';
  WARN_DNV_PMIN_GT_DESIGN_PRESS = 'Warning minimum pressure for DNV greater than design pressure for element: ';
  WARN_DNV_ACT_PRESS_GT_DESIGN_PRESS = ' Warning actual pressure greater than design pressure for DNV code';
  WARN_DNV_PIPE_BURST = ' Warning: Pipe has burst on pressure alone for element ';
  WARN_ISO_YLD_GT_UTS = ' Warning: Yield Strength larger than Ultimate Tensile Strength';
  WARN_ISO_PIPE_BURST = ' Warning: Pipe burst';
  WARN_ISO_OVALITY_MIN = ' Warning: Ovality overwritten - minimum initial ovality 0.25%';
  WARN_ISO_OVALITY_MAX = ' Warning: Initial ovality for deep water should generally not exceed 0.5%';
  WARN_ISO_TE_NEG = 'Warning: Effective Tension Negative in ISO Check, ISO Check Valid for Elements in Tension Only';
  WARN_DNV_TE_NEG = 'Warning: Effective Tension Negative in DNV Check, DNV Check Valid for Elements in Tension Only';
  NOTE_API_STD_2RD_M_MAX = 'BM/0.9Mmax > 1.0, use nonlinear stress-strain properties';
  NOTE_API_STD_2RD_BSTRAIN = 'Bending Strain/0.5% > 1.0, use strain based design method';
  WARN_API_STD_2RD_TCOUTBOUNDS = 'Warning: Out of bounds values in thermal curve in set:';
  WARN_SSCURVE = 'Warning: Old keyword SSCURVE still present in INF File';
  WARN_BSI_STT = 'Warning: For time domain BSI code check no timetrace of Shear loading used';
  WARN_PIPE_SLENDERNESS = 'Warning: Pipe Slenderness out of Range in ISO Check, Setting to Alpha Bm to 0.001 for affected elements';
  WARN_PIPE_PRESSURE = 'Warning: Pipe Pressure is significantly different from Initial Configuration - 2HSTRE3D only uses initial pressure for element : ';

type
  T2DArray = Array of Array of Real;
  APISTD2RDArray = array [0..9] of Real;

var
  WARN_ISO_TE_NEG_FLAG, WARN_DNV_TE_NEG_FLAG, WARN_BSI_STTFLAG : boolean;
  WARN_PIPE_SLENDERNESSFLAG, WARN_PIPE_PRESSUREFLAG   : boolean;
  APISTDmethod4, BSIShearInput                                : Boolean;
  ExtremeData                                                 : Boolean;
  longStr                                                     : String[36];
  filename, FilenameF, infFile,
  FilenameF2, tempStr, dumStr, tempDumpString                 : String;
  title                                                       : String[150];
  limitState                                                  : String;
  b, o, t, h, x,g                                          : Text;
  Program_Filename, CurrentDirProgram                         : String;
  lastparam                                                   : string;
  ParamCountRemain                                            : integer;

  i, totalNodes, totalElements, totalSprings, MaxFluidVal,
  totalLinFlexJoints, totalNonlinFlexJoints, totalFlexJoints,
  code, flag,
  Num_M_Max, Num_Bending_Strain, Num_Temperature,
  fluidSetNo, fluidVal,
  flexVer1, flexVer2, flexVer3,
  currentFileNo, maxFileNo, rRef, warningsCount               : Integer;

  Delta0,
  YieldStr_Fac, TensileStr_Fac,
  KFactor, nFactor, StressVal, SSslope,
  FluidHeight, FluidDens, FluidPress, PressCorr,
  SuppElev, TubZeroTens, IntCoat                                              : Real;

  debug, // Variable to print help and errors checking to screen
  APIRP2RDCodeCheck, APISTD2RDCodeCheck, DNVCodeCheck, STT, OSTT,
  PRINTTB2, PRINTTB1, pWarn,pWarn2, DNVPinput, DNV_Code_Mod ,
  ExtFluidIndex, IntFluidIndex, // ExtFluidIndex and IntFluidIndex flag if internal and external fluid properties are included in the INF File
  ISOCodeCheck, BSICodeCheck, // ISO , BSI criterion

  angleTablesPresent, warnFlag1, warnTemp, coatOdflag         : Boolean;
  globalFactor, derateFactor, Functional,OutputToExcel,
  IntFluidElementOverride  : Boolean;
  INFnameInOutput : Boolean;

  screenX, screenY                                            : Short;

  // Dynamic arrays
  flexJoint                   : Array of Array of Integer;
  linFJStiffness              : Array of Real;
  headerFJArr                 : Array of String;
  outputFJArr                 : Array of Array of String;
  OSTTArray                   : Array of Array of Array of real;

  nonlinFJTableMom,
  nonlinFJTableAng            : T2DArray;

  TraceString                 : Array[1..MAX_ELEMENTS,1..5,1..2] of integer;
  TraceRead                   : Array[1..MAX_TIMETRACES] of Real;
  warningsArr                 : Array[1..MAX_ELEMENTS] of String;
  APISTDValMethod4            : Array[1..MAX_ELEMENTS] of Boolean;
  spring                      : Array[1..MAX_SPRINGS] of Integer;
  fluid, intFluid             : Array[0..MAX_FLUIDS,1..3] of Real;
  extFluid                    : Array[1..MAX_ELEMENTS,1..3] of Real;
  intFluidElement             : Array[1..MAX_ELEMENTS,1..3] of Real;
  STTData                     : Array[1..MAX_ELEMENTS,1..9] of Real;
  tmax, tmin, mmax, mmin      : Array[1..MAX_ELEMENTS] of Real;
  tmaxSTD, tminSTD            : Array[1..MAX_ELEMENTS] of Real;
  mmaxSTD, mminSTD            : Array[1..MAX_ELEMENTS] of Real;
  STTElem                     : Array[1..MAX_ELEMENTS] of Real;
  DNVFactors                  : Array[1..MAX_ELEMENTS,1..7] of Real;
  DNVMatSpec                  : Array[1..MAX_ELEMENTS,1..6] of Real;
  DNVMults                    : Array[1..MAX_ELEMENTS,1..2] of Real;
  DNVPres                     : Array[1..MAX_ELEMENTS,1..6] of Real;
  DNVInt, DNVExt, DNVmax,
  DNVIntMax, DNVExtMax        : Array[1..MAX_ELEMENTS,1..2] of Real;
  elemSwitch                  : Array[1..MAX_ELEMENTS] of Real;
  temperature                 : Array[1..MAX_ELEMENTS,1..2] of Real;
  SAFactor                    : Array[1..MAX_ELEMENTS] of Real;
  BurstK                      : Array[1..MAX_ELEMENTS] of Real;
  Alphafab                    : Array[1..MAX_ELEMENTS] of Real;
  FcM4                        : Array[1..MAX_ELEMENTS] of Real;
  Collapse_Type               : Array[1..MAX_ELEMENTS] of Boolean;
  temperatureFactor           : Array[1..MAX_ELEMENTS] of Real;
  temperatureFactorTensile    : Array[1..MAX_ELEMENTS] of Real;
  ElemInGRD                   : Array[1..MAX_ELEMENTS] of Integer;

  // Labels arrays
  Labeldata : boolean;
  LabelElem, LabelNode        : Array[1..MAX_LABELS] of string;
  LabelElemI, LabelNodeI      : Array[1..MAX_LABELS,1..2] of integer;
  TotalLabelE, TotalLabelN    : Integer;
  TotalSingleElemLabels       : Integer;

  // Stress Check output array
  StressCheckA                : Array[1..MAX_ELEMENTS,1..2,1..11] of real;
  STD2RDOut                   : Array[1..MAX_ELEMENTS,1..10] of real;
  //----------------
  // Loading inputs
  //----------------
  loads                       : Array[1..MAX_ELEMENTS,1..18,1..2,1..3] of Real;
  PInfo                       : Array[1..MAX_ELEMENTS,1..12] of Real;
  Lheader                     : Array[1..3] of String;
  // Connectivity matrix
  elemInfo                    : Array[1..MAX_ELEMENTS,1..4] of Integer;
  //  10 is Ultimate Strength, 11 is Young's Modulus, 12 is Ovality, 13 is Poisson's ratio
  pipeInfo                    : Array[1..MAX_ELEMENTS,1..14] of Real;
  nodeinfoI                   : Array[1..MAX_NODES] of Integer;
  nodeInfo                    : Array[1..MAX_NODES,1..15] of real;
  springdata                  : Array[1..MAX_SPRINGS,1..7] of real;
  springMatdata               : Array[1..MAX_SPRINGS,1..MAX_MATERIAL_POINTS,1..2] of real;
  // This array stores nominal static tension values from inf file
  StatTen                     : Array[1..MAX_ELEMENTS] of Real;
  // This array 'OSTTElem is used when the keyword 'OSTT' is used. It flags elements for
  // which stress timetraces need to be output
  OSTTElem                    : Array[1..MAX_ELEMENTS] of Integer;
  OSTTData                    : Array[1..MAX_ELEMENTS] of Integer;
  OSTTNoElementStressTime, OSTTTotalNoTimeStep : integer;
  zeroPressureCheck           : Array[1..MAX_ELEMENTS] of Integer;
  // Tension correction
  tensCorr                    : Array[1..MAX_ELEMENTS] of Real;
  factors                     : Array[1..MAX_ELEMENTS,1..7] of Real;
  YieldStrTempCor             : Array[1..10,0..20,1..2] of Real;

  tempel : integer;
  tempstrax, tempstrbmsign, tempstrhoop,
  tempstrrad,tempshrcomb , tempvmout     : real;
  Maxstrax, Maxstrbmsign, Maxstrhoop,
  Maxstrrad, Maxshrcomb, Maxvmout        : array [1..MAX_ELEMENTS,1..2] of real;

  HeaderToExcel, DataSToExcel, DataS2ToExcel : array of array of olevariant;
  DataToExcel, Data2ToExcel : array of array of olevariant;

  // Array for STT loads 
  OSTTiind : array[1..MAX_ELEMENTS] of integer;
  stressmax : Array[1..MAX_ELEMENTS,1..5] of Real;
  count : Array[1..MAX_ELEMENTS,1..3] of Integer;

{ **************************************************************************** }
// Writes file headers.
{ **************************************************************************** }
procedure Header(pageTitle : String);

var
    StressCheckString : string;

begin
  WriteLn(b, 'Program : ',PROG_NAME,'       by 2H Offshore Engineering Limited');
  WriteLn(b, 'Version : ', VERSION);

  if DNVCodeCheck then
    begin
      if MaxfileNo = 1 then
        WriteLn(b, 'Functional Filename   : ', filename);

      if MaxfileNo = 2 then
        begin;
        WriteLn(b, 'Enviromental Filename : ', filename);
        Writeln(b, 'Functional Filename   : ', filenameF);
        end;

      if MaxfileNo = 3 then
        begin;
        WriteLn(b, 'Enviromental Filename  : ', filename);
        Writeln(b, 'Functional Filename    : ', filenameF);
        Writeln(b, 'Enviromental2 Filename : ', filenameF2);
        end;
    end

  else

  Writeln(b, 'Filename : ', filename);

  Writeln(b,'.INF File : ', infFile);

  StressCheckString := 'BSI';
  if ISOCodeCheck = True then StressCheckString := 'ISO';  // Stress Check Type
  if APIRP2RDCodeCheck = True then StressCheckString := 'API-RP-2RD';
  if DNVCodeCheck = True then StressCheckString := 'DNV-OS-F201';
  if APISTD2RDCodeCheck = True then StressCheckString := 'API-STD-2RD';

  Writeln(b,' Stress Check to :', StressCheckString);

  if flexVer1 >= 6 then
    WriteLn(b, 'Title   : ', Copy(title, 2, 78))
  else
    WriteLn(b, 'Title   : ', Copy(title, 3, 78));

  WriteLn(b,'                                               ', pageTitle);
  WriteLn(b);
end;  // Header


{ **************************************************************************** }
// Ouput guidance on running program, e.g. necessary parameters.
{ **************************************************************************** }
procedure Help;

begin
  // Tell the user what's wrong!
  WriteLn (' Not enough command line parameters - atleast 2 are required');
  WriteLn;
  WriteLn(' For API and BSI check');
  WriteLn (' 2HSTRE3D [Run Name] [*.INF File Name]');
  WriteLn;
  WriteLn(' For DNV Check');
  WriteLn(' 2HSTRE3D [Run Name] [*.INF File Name] [Optional Functional Loading File]');
  Halt(1);
end;  // Help


{ **************************************************************************** }
// Ensure the .tab file is correctly assigned.
// The .tab file may have a "-pst" suffix in versions 8 and greater.
{ **************************************************************************** }
procedure AssignTabFile(filename : String);

begin
  if FileSearch(filename + '.tab','') <> '' then
    AssignFile(t, filename + '.tab')
  else
    if FileSearch(filename + '-pst.tab','') <> '' then AssignFile(t, filename + '-pst.tab')
    else
    begin;
    Writeln(' No tab file no found'); halt(1);
    end;
  Reset(t);
end;  // AssignTabFile


{ **************************************************************************** }
// Finds the Flexcom version number in the .out file and
// returns the first and second version index numbers.
{ **************************************************************************** }
procedure GetFlexcomVersion(var searchFile : TextFile; out flexVer1, flexVer2 : Integer);

var
  fileLine, verNo : String;
  verPos, offset, verLength : Integer;
  found : Boolean;

begin
  // Read from beginning of file
  Reset(searchFile);

  // Version number found flag
  found := false;
  // Number of characters from start of 'Version' to version number
  offset := 8;
  // Length of version number
  verLength := 5;

  // Search each line of file until line containing version number is found
  while not EoF(searchFile) and not found do begin
    ReadLn(searchFile, fileLine);
    fileLine := LowerCase(fileLine);

    if AnsiContainsStr(fileLine, 'version') then found := True;
  end;

  // Stop program if version number not found
  if EoF(searchFile) then begin
      WriteLn(' Error: Flexcom version number not found in ', filename, '.out file!');
      Halt(1);
  end;

  // Determine location of 'Version' in line
  verPos := Pos('version', fileLine);

  // Extract version number
  verNo := Copy(fileLine, verPos + offset, verLength);

  // Extract first and second components of version number
  flexVer1 := StrToInt(Copy(verNo,1,1));
  flexVer2 := StrToInt(Copy(verNo,3,1));
  flexVer3 := StrToInt(Copy(verNo,5,1));

  if (flexVer1 = 8) and (flexVer2 = 3) and (flexVer3 = 7) then
    begin;
    writeln('Error: 2HSTRE3D will not work for version 8.3.7 of Flexcom');
    halt(1);
    end;
end;  // GetFlexcomVersion

{ **************************************************************************** }
// Calculate  Pc of Equ. 5 in API-RP-2RD
{ **************************************************************************** }
function GetPcFromCubicEqu(PpVal,PelVal,DVal,tVal,delta0 : Real) : Real;
Var
    b, c, d, p, u, v, y    : Real;
begin
    b := -1*PelVal;
    c := -1*(sqr(PpVal) + PelVal*PpVal*2*Delta0*DVal/tVal);
    d := PelVal*sqr(PpVal);

    u := 1/3*(c - 1/3*sqr(b));
    v := 1/2*(2/27*power(b,3.0)-1/3*b*c + d);
    p := arccos(-1*v/power(-1*u,3/2));
    y := -2*sqrt(-1*u)*cos(p/3+pi/3);

    result := y-1/3*b;
end;  //GetPcFromCubicEqu

{ MAIN API-STD-2RD Calculation rountine }

function APISTD2RDCalc(piVal,peVal,PbVal,PcVal,MpVal,TyVal,MyVal,TenVal,MVal,CurVal,yieldStrVal,IVal,EVal,ODVal,IDVal: Real; elemNum: Integer) : APISTD2RDArray;
var
    M_max, FDVal, FDVal1, FDVal2, FDValPre, temp, tempRT, bendingStr  : Real;
    epsilonVal : real;
    tempStep : Real;

{ Output data in APISTD2RDArray as follows,
   [1]  Bending Moment Ratio test
   [2]  Method 1 Utilisation
   [3]  Method 2 Utilisation
   [4]  Method 4 Utilisation (pressure and tension)
   [5]  Method 4 Bending Strain Criteria
   [6]  Bending Strain Applicable for method 4 
   Note method 3 Utilisation is DNV calculation }

begin
    epsilonVal := (ODVal-IDVal)/4.0/ODVal;                                        //Equ. 9 in API STR 2RD
    EVal := EVal * power(10,6); { Convert from MPa to Pa }
    M_max := 0;
    Delta0 := pipeInfo[i,12];
    { determination of Bending Moment capacity for method 4. }
    if PiVal >= PeVal then  { internal Overpressure }
    if ((1-sqr((piVal-peVal)/PbVal)) > 0) then
      if  abs(TenVal/TyVal/sqrt(1-sqr((piVal-peVal)/PbVal))) < 1 then
         M_max := MpVal*sqrt(1.0-sqr((piVal-peVal)/PbVal)) * cos(pi/2.0*TenVal/TyVal/
                            sqrt(1-sqr((piVal-peVal)/PbVal)));  // Equ. 30 in API-STD-2RD
    if PeVal > PiVal then   { external Overpressure }
      if ((1-sqr((peVal-piVal)/PcVal)) > 0) then
      if  abs(TenVal/TyVal/sqrt(1-sqr((peVal-piVal)/PcVal))) < 1 then
         M_max := MpVal*sqrt(1.0-sqr((peVal-piVal)/PcVal)) * cos(pi/2.0*TenVal/TyVal/
                            sqrt(1-sqr((peVal-piVal)/PcVal)));  // Equ. 30 in API-STD-2RD (modified external overpressure)
    if M_Max > 1E-10 then result[0] := MVal/0.9/M_max
      else result[0] := 1E10;

    if SAFactor[elemNum] > 0 then
    begin;
    bendingStr := CurVal*ODVal/2 * SAFactor[elemNum];
    end
    else
    bendingStr := -1;
    result[5] := bendingstr;

    if bendingStr > 0.005 then  // Note on above max bending strain > 0.5%
        inc(Num_Bending_Strain,1);

    if MVal > 0.9*M_max then   // Note Check Bending Strain
        inc(Num_M_Max,1);


    if (piVal > peVal) then //internal over pressure
    begin
        result[1] := sqrt(sqr(abs(TenVal/TyVal)+abs(MVal/MyVal))+sqr((piVal-peVal)/PbVal)); //Method 1 Equ. 18 in API-STD-2RD
        FDVal1 := result[1];
        FDVal2 := result[1];
        if sqr(FDVal1)-sqr((piVal-peVal)/PbVal) < 0 then
        begin
            FDVal1 := abs((piVal-peVal)/PbVal)*1.001;
            FDVal2 := FDVal1;
        end;
        tempRT := sqrt(sqr(FDVal1)-sqr((piVal-peVal)/PbVal));
        if (abs(TenVal/TyVal/TempRT) > 1) then
        temp := - abs(MVal/MpVal)
        else
        temp := tempRT*cos(pi/2*abs(TenVal/TyVal/tempRT)) - abs(MVal/MpVal); // Method 2 Equ. 21 in API-STD-2RD
        tempStep := 0.1*FDVAl1;
        if temp > 0 then  //Recalculate the lower bound
        begin
            while (temp > 0) and (FDVal1 >= abs((piVal-peVal)/PbVal)) do
            begin
                FDVal1 := FDVal1 - tempStep;
                if FDVal1 <abs((piVal-peVal)/PbVal) then
                begin
                    FDVal1 := abs((piVal-peVal)/PbVal);
                    break;
                end;
                tempRT := sqrt(sqr(FDVal1)-sqr((piVal-peVal)/PbVal));
                if (abs(TenVal/TyVal/TempRT) > 1) then
                temp := - abs(MVal/MpVal)
                else
                temp := tempRT*cos(pi/2*abs(TenVal/TyVal/tempRT)) - abs(MVal/MpVal); // Method 2 Equ. 21 in API-STD-2RD
            end;
        end
        else
        begin              //Recalculate the upper bound
            while (temp < 0) do
            begin
                FDVal2 := FDVal2 + tempStep;
                tempRT := sqrt(sqr(FDVal2)-sqr((piVal-peVal)/PbVal));
                if (abs(TenVal/TyVal/TempRT) > 1) then
                temp := - abs(MVal/MpVal)
                else
                temp := tempRT*cos(pi/2*abs(TenVal/TyVal/tempRT)) - abs(MVal/MpVal); // Method 2 Equ. 21 in API-STD-2RD
            end;
        end;
        FDValPre := FDVal1;
        FDVal := FDVal2;
        while abs(FDVal - FDValPre) > 0.0005 do
        begin
            tempRT := sqrt(sqr(FDVal)-sqr((piVal-peVal)/PbVal));
            if (abs(TenVal/TyVal/TempRT) > 1) then
                temp := - abs(MVal/MpVal)
                else
            temp := tempRT*cos(pi/2*abs(TenVal/TyVal/tempRT)) - abs(MVal/MpVal); // Method 2 Equ. 21 in API-STD-2RD
            if temp >= 0 then
                FDVal2 := FDVal
            else
                FDVal1 := FDVal;
            FDValPre := FDVal;
            FDVal := (FDVal2 + FDVal1)*0.5
        end;
        result[2] := FDVal;
        result[3] := sqrt(sqr((piVal-peVal)/PbVal)+sqr(TenVal/TyVal));  // Method 4 Tension Equ. 24 in API-STD-2RD
        result[4] := bendingStr/epsilonVal; // Method 4 Equ. 26 in API-STD-2RD
    end
    else //external overpressure
    begin
        result[1] := sqrt(sqr(abs(TenVal/TyVal)+abs(MVal/MyVal))+sqr((peVal-piVal)/PcVal));//Method 1 Equ. 19 in API-STD-2RD
        FDVal1 := result[1];
        FDVal2 := result[1];
        if sqr(FDVal1)-sqr((peVal-piVal)/PcVal) < 0 then
        begin
            FDVal1 := abs((peVal-piVal)/PcVal)*1.001;
            FDVal2 := FDVal1;
        end;
        tempRT := sqrt(sqr(FDVal1)-sqr((peVal-piVal)/PcVal));
        if (abs(TenVal/TyVal/TempRT) > 1) then
        temp := - abs(MVal/MpVal)
        else
        temp := tempRT*cos(pi/2*abs(TenVal/TyVal/tempRT)) - abs(MVal/MpVal); // Method 2 Equ. 22 in API-STD-2RD
        tempStep := 0.1*FDVAl1;
        if temp > 0 then  //Recalculate the lower bound
        begin
            while (temp > 0) and (FDVal1 >= abs((peVal-piVal)/PcVal)) do
            begin
                FDVal1 := FDVal1 - tempStep;
                if FDVal1 < abs((peVal-piVal)/PcVal) then
                begin
                    FDVal1 := abs((peVal-piVal)/PcVal)+ 0.00001;
                    break;
                end;
                tempRT := sqrt(sqr(FDVal1)-sqr((peVal-piVal)/PcVal));
                if (TenVal/TyVal/TempRT > 1) then
                temp := - abs(MVal/MpVal)
                else
                temp := tempRT*cos(pi/2*abs(TenVal/TyVal/tempRT)) - abs(MVal/MpVal); // Method 2 Equ. 22 in API-STD-2RD
            end;
        end
        else
        begin              //Recalculate the upper bound
            while (temp < 0) do
            begin
                FDVal2 := FDVal2 + tempStep;
                tempRT := sqrt(sqr(FDVal2)-sqr((peVal-piVal)/PcVal));
                if (abs(TenVal/TyVal/TempRT) > 1) then
                temp := - abs(MVal/MpVal)
                else
                temp := tempRT*cos(pi/2*abs(TenVal/TyVal/tempRT)) - abs(MVal/MpVal); // Method 2 Equ. 22 in API-STD-2RD
            end;
        end;
        FDValPre := FDVal1;
        FDVal := FDVal2;
        while abs(FDVal - FDValPre) > 0.0005 do
        begin
            tempRT := sqrt(sqr(FDVal)-sqr((peVal-piVal)/PcVal));
            if (TenVal/TyVal/TempRT > 1) then
                temp := - abs(MVal/MpVal)
            else
            temp := tempRT*cos(pi/2*abs(TenVal/TyVal/tempRT)) - abs(MVal/MpVal); // Method 2 Equ. 22 in API-STD-2RD
            if temp >= 0 then
                FDVal2 := FDVal
            else
                FDVal1 := FDVal;
            FDValPre := FDVal;
            FDVal := (FDVal2 + FDVal1)*0.5
        end;
        result[2] := FDVal;
        result[3] := sqrt(sqr((peVal-PiVal)/PcVal) + sqr(TenVal/TyVal));                        // Method 4 Tension Equ. 24 in API-STD-2RD with Pc.
        result[4] := bendingStr/epsilonVal*(1+20*Delta0)/(1-(peVal-piVal)/FcM4[elemNum]/PcVal); // Method 4 Equ. 27 in API-STD-2RD
    end;
    // readjust for limit states.
    if limitState = 'SLS' then
    begin
        result[1] := result[1]/0.8; //Equ. 20 in API-STD-2RD
        result[2] := result[2]/0.8; //Equ. 23 in API-STD-2RD
        result[3] := result[3]/0.9; //Equ. 25 in API-STD-2RD
        result[4] := result[4]/0.5; //Equ. 28 in API-STD-2RD
    end
    else if limitState = 'ULS' then
    begin
        result[1] := result[1]/0.8; //Equ. 20 in API-STD-2RD
        result[2] := result[2]/0.8; //Equ. 23 in API-STD-2RD
        result[3] := result[3]/0.9; //Equ. 25 in API-STD-2RD
        result[4] := result[4]/0.5; //Equ. 28 in API-STD-2RD
    end
    else if limitState = 'ALS' then
    begin
        if (peVal-piVal) > 0.5*PcVal then  //Equ. 20 in API-STD-2RD
            result[2] := result[2]/0.9;
    end;
    if SAFactor[elemNum] < 0 then result[4] := -1;

end;   //APISTD2RDCalc

{ **************************************************************************** }
// 1.51 ISO function.
{ **************************************************************************** }
procedure ISOCWO(Te, Mbm, PInt, PExt, Dout, Thickness, Yield, ultYield,
  EModulus, Ovality, Poisson : Real; var ISOUtilizationNumber : Array of Real);
{
This Function calculates and returns the Utilization given in Eq. 17 or Eq. 25, Section 6.5 of ISO 13628-7:2005
The formulation given by Eq. 17 assumes internal pressure is >= external pressure
The formulation given by Eq. 25 assumes internal pressure is lower than external pressure

Input:
Te is the effective tension in the pipe;
Mbm is the bending moment in the pipe;
PInt is the internal pressure in the pipe;
PExt is the external pressure;
Dout is the specified or nominal pipe outside diameter;
Thickness is the pipe wall thickness without allowance (norminal t minus corrosion allowance); see 6.5.3.1;
Yield is the design yield strength; see 6.4.6;
ultYield is the design ultimate tensile strength; see 6.4.6.
EModulus is the Young's modulus of the pipe
Ovality is the initial ovality
Poisson is the Poisson's ratio

Important Intermediate Variables:
Alphabm is the pipe cross-section slenderness parameter;
Tpc is the plastic tension capacity of the pipe;
Mpc is the plastic bending moment capacity of the pipe;
Pb is the burst pressure of the pipe;
Pc is the pipe hoop buckling/collapse pressure
Pel is the elastic hoop buckling/collapse pressure
Pp is the plastic pressure at hoop bukling

Output:
The value of the LHS of Eq. 17 or 25
ISOUtilizationNumber updates if the new values are larger
}

var
  i : Integer;
  Result, Alphabm, Tpc, Mpc, Pb, SigyDoET2, Pc, Pel, Pp : Real;
  Fd : array[0..3] of Real;
  SolveB, SolveC, SolveD, SolveU, SolveV, Solve2, SolvePhi : Real;

begin
  ISOUtilizationNumber[0]:=0;
  ISOUtilizationNumber[1]:=0;
  ISOUtilizationNumber[2]:=0;
  ISOUtilizationNumber[3]:=0;
  if Yield > ultYield then begin
    WriteLn(WARN_ISO_YLD_GT_UTS);
    Inc(warningsCount);
    warningsArr[warningsCount] := WARN_ISO_YLD_GT_UTS;
  end;
  if (Te < 0) and (WARN_ISO_TE_NEG_FLAG = false) then begin;
    Writeln(WARN_ISO_TE_NEG);
    WARN_ISO_TE_NEG_FLAG := true;
    Inc(warningsCount);
    warningsArr[warningsCount] := WARN_ISO_TE_NEG;
    end;

  Fd[0] := 0.67;
  Fd[1] := 0.80;
  Fd[2] := 0.90;
  Fd[3] := 1.00;

  // Mpc and Tpc
  SigyDoET2 := Yield*Dout/(Emodulus*Thickness);

  if (SigyDoET2<=0.0517) then
    Alphabm := 1.0
  else if SigyDoET2 <= 0.1034 then
    Alphabm := 1.13 - 2.58*SigyDoET2
  else if SigyDoET2 <= 0.170 then
    Alphabm := 0.94 - 0.76*SigyDoET2
  else
    begin
      Alphabm := 0.001;
      if WARN_PIPE_SLENDERNESSFlag = false then
       begin;
       WriteLn('*** Warning, Pipe Slenderness out of Range in ISO Check, Setting to Alpha Bm to 0.001');
       Inc(warningsCount);
       warningsArr[warningsCount] := WARN_PIPE_SLENDERNESS;
       Warn_Pipe_SlendernessFlag := true;
       end;
    end;

  Mpc := Alphabm*Yield/6*(Power(Dout, 3) - Power(Dout - 2*Thickness, 3));
  Tpc := Yield*3.1415926535*(Dout - Thickness)*Thickness;

  // Internal overpressure
  if PInt >= PExt then
    begin
      Pb := 1.1*(Yield + ultYield)*Thickness/(Dout - Thickness);

      for i := Low(ISOUtilizationNumber) to High(ISOUtilizationNumber) do begin

        Result := Power(Te/(Fd[i]*Tpc), 2) + Power((PInt - PExt)/(Fd[i]*Pb), 2);

        if (PInt - PExt)/(Fd[i]*Pb) <= 1 then
          Result := Result + abs(Mbm)*Sqrt(1 - Power((PInt-PExt)/(Fd[i]*Pb), 2))/(Fd[i]*Mpc)
        else
          begin
            WriteLn(WARN_ISO_PIPE_BURST);
            Inc(warningsCount);
            warningsArr[warningsCount] := WARN_ISO_PIPE_BURST;
          end;

        ISOUtilizationNumber[i]:=max(ISOUtilizationNumber[i],Result);

      end;
    end

  // External overpressure
  else
    begin
      // Checking ovality: ISO CODE specifies 0.0025 <= ovality <= 0.015, and for deep water ovality <= 0.005
      if Ovality < 0.0025 then
        begin
          Ovality := 0.0025;
          WriteLn(WARN_ISO_OVALITY_MIN);
          Inc(warningsCount);
          warningsArr[warningsCount] := WARN_ISO_OVALITY_MIN;
        end

      else if Ovality > 0.005 then
        begin
          WriteLn(WARN_ISO_OVALITY_MAX);
          Inc(warningsCount);
          warningsArr[warningsCount] := WARN_ISO_OVALITY_MAX;
        end;

      // Solving for Pc
      Pel := 2*EModulus*Power(Thickness/(Dout - Thickness), 3)/(1 - Poisson*Poisson);
      Pp := 2*Yield*Thickness/Dout;

      // Algorithm in Annex E of ISO code. Same as those in DNV check except ovality definition
      SolveB := -Pel;
      SolveC := -(Sqr(Pp) + 2*Pp*Pel*Ovality*Dout/Thickness);
      SolveD := Pel*Sqr(Pp);
      SolveU := 1/3*(-1/3*Sqr(SolveB) + SolveC);
      SolveV := 1/2*(2/27*power(SolveB, 3) - 1/3*SolveB*SolveC + SolveD);
      Solve2 := -SolveV/Sqrt(-power(SolveU, 3));

      if (Solve2 >= -1) And (Solve2 < 1) then
        SolvePhi := arccos(Solve2)
      else
        begin
          WriteLn(' Error - Has occured in pressure collaspe solver');
          if debug then WriteLn(' Solve2 is ', Solve2);
          Halt(2);
        end;

        pc := -2*Sqrt(-SolveU)*cos(1/3*SolvePhi + 1/3*pi) - 1/3*SolveB;

        // Calculating Eq 25
        for i := low(ISOUtilizationNumber) to high(ISOUtilizationNumber) do begin
          Result:=Sqr(Sqr(Te/Fd[i]/Tpc)+Mbm/0.95/Fd[i]/Mpc)+Sqr((PExt-PInt)/Fd[i]/Pc);
          ISOUtilizationNumber[i]:=max(ISOUtilizationNumber[i],Result);
        end;

    end;

end;  // ISOCWO

{ **************************************************************************** }
// Calculate yield strength temperature correction factor.
{ **************************************************************************** }
procedure GetYieldStrFac(elemN : integer);

var
    i, numOfPoints,counter,setNum: Integer;

begin
    setNum := round(temperature[elemN,2]);
    if setNum < 1 then begin; writeln('ERROR in Specfication of Thermal Table or Temperature Value'); halt(1); end;
    if setNum > 10 then begin; writeln('Maximum Thermal Table Reached '); halt(1); end;
    numOfPoints := round(YieldStrTempCor[setNum,0,1]);// YieldStrTempCor[setNum,0,1] is the number of points on temperature set "setNum"
    counter := 1;
    for i := 1 to numOfPoints - 1 do
    begin
        if temperature[elemN,1] < YieldStrTempCor[setNum,1,1] then
        begin
            YieldStr_Fac := YieldStrTempCor[setNum,1,2];
            // warning that out of bounds on data ?
            if YieldStrTempCor[setNum,0,2] <> 1 then
             begin;
             inc(warningscount);
             Warningsarr[warningsCount] := WARN_API_STD_2RD_TCOUTBOUNDS + InttoStr(SetNum);
             YieldStrTempCor[SetNum,0,2]:= 1;
             end;
            break;
        end
        else if (temperature[elemN,1] >= YieldStrTempCor[setNum,i,1]) and (temperature[elemN,1] <= YieldStrTempCor[setNum,i+1,1]) then
        begin
            YieldStr_Fac := YieldStrTempCor[setNum,i,2] + (YieldStrTempCor[setNum,i+1,2]-YieldStrTempCor[setNum,i,2])/
            (YieldStrTempCor[setNum,i+1,1]-YieldStrTempCor[setNum,i,1])*(temperature[elemN,1]-YieldStrTempCor[setNum,i,1]);
            break;
        end;
        counter := counter + 1;
    end;
    if counter >= numOfPoints then
        begin;
        YieldStr_Fac := YieldStrTempCor[setNum,numOfPoints,2];
        // Warning that out of bounds on data
        if YieldStrTempCor[setNum,0,2] <> 1 then
             begin;
             inc(warningscount);
             Warningsarr[warningsCount] := WARN_API_STD_2RD_TCOUTBOUNDS  + InttoStr(SetNum);
             YieldStrTempCor[SetNum,0,2]:= 1;
             end;
        end;
    TensileStr_Fac := YieldStr_Fac;
end;

{ **************************************************************************** }
// Read model discretisation and setup details from .out file.
{ **************************************************************************** }
procedure GetModelDetails;

var
  temp, dumStr1 : String;
  count, count2, code, elementNo, matlNo, loopCounter, i, j, k : Integer;
  stiffness : Real;
  warnSpring : Boolean;
  linFJElem, nonlinFJElem : Array of Integer;

begin
  // Initialise dynamic arrays
  SetLength(linFJElem, MAX_SPRINGS);
  SetLength(nonlinFJElem, MAX_SPRINGS);
  SetLength(flexJoint, MAX_SPRINGS, 3);
  SetLength(linFJStiffness, MAX_SPRINGS);
  SetLength(nonlinFJTableMom, MAX_SPRINGS, MAX_SPRINGS);
  SetLength(nonlinFJTableAng, MAX_SPRINGS, MAX_SPRINGS);

  warnSpring := False;

  //--------------------
  // Number of elements
  //--------------------
  if flexVer1 >= 5 then
    begin
      dumStr := FindLine(o, 'Total No. of Structural Elements');
      FirstWord(dumStr, temp);
    end
  else
    dumStr := FindLine(o, 'Total No. of Elements');

  // Remove the text at the begining of the dumStr line
  for i := 1 to 6 do FirstWord(dumStr, temp);

  // Store total number of elements as integer
  totalElements := StrToInt(temp);

  if totalElements > MAX_ELEMENTS then begin
    WriteLn('Number of elements is ', totalElements, ' maximum number allowed :', MAX_ELEMENTS);
    Halt(1);
  end;

  //-----------------
  // Number of nodes
  //-----------------
  dumStr := FindLine(o, 'No. of Nodes');

  // Remove the text at the beginning of the dumStr line
  for i := 1 to 5 do FirstWord(dumStr, temp);

  if debug = True then WriteLn(' Nodes: ', temp);
  // Converts string to number
  totalNodes := StrToInt(temp);

  //--------------------
  // Linear flex joints
  //--------------------
  dumStr := FindLine(o, 'No. of Linear Flex Joints');

  // Remove the text at the begining of the dumStr line
  for i := 1 to 7 do FirstWord(dumStr, temp);
  Val(temp, totalLinFlexJoints, code);

  if debug = True then WriteLn (' Number of Linear Flex Joints: ', temp);

  // If linear flex joints exist, extract data
  if totalLinFlexJoints > 0 then begin

    // Find Linear Flex Joints table
    dumStr := FindLine(o, '*** LINEAR FLEX JOINT DATA ***');

    for loopCounter := 1 to 4 do ReadLn(o, dumStr);
    if pos('Element',dumstr) > 0 then Readln(o,dumstr);

    // Store linear flex joint information to arrays
    // (note using dynamic array, which starts at index 0)
    count := 0;
    while count < totalLinFlexJoints do begin
      // Read the element number
      ReadLn(o, temp);
      FirstWord(temp, dumStr);

      // Store linear flex joint element in array
      Val(dumStr, linFJElem[count], code);

      // And the rest of the data
      FirstWord(temp, dumStr);               // Weight in Air
      FirstWord(temp, dumStr);               // Weight in Water
      FirstWord(temp, dumStr);               // Stiffness

      // Store linear flex joint stiffness in array
      Val(dumStr, linFJStiffness[count], code);

      Inc(count);

      // Check number of linear flex joints in file does not exceed storage limit
      if count > MAX_SPRINGS then begin
        WriteLn(' Error: There are ', count, ' linear flex joints in the model - this is too many for ', PROG_NAME);
        Halt(1);
      end;

    end;

    // Set size of array
    SetLength(linFJElem, totalLinFlexJoints+1);
    SetLength(linFJStiffness, totalLinFlexJoints+1);

  end;

  //-----------------------
  // Nonlinear flex joints
  //-----------------------
  // Determine number of nonlinear flex joints
  dumStr := FindLine(o, 'No. of Non-Linear Flex Joint');

  // Remove the text at the begining of the dumStr line
  for i := 1 to 7 do FirstWord(dumStr, temp);
  Val(temp, totalNonlinFlexJoints, code);

  if debug = True then WriteLn(' Number of Nonlinear Flex Joints: ', temp);

  // If nonlinear flex joints exist, extract data
  if totalNonlinFlexJoints > 0 then begin

    // Find Nonlinear Flex Joints table
    dumStr := FindLine(o, '*** NONLINEAR FLEX JOINT DATA ***');

    // Skip to line containing element data
    for loopCounter := 1 to 4 do ReadLn(o, dumStr);

    // Store nonlinear flex joint elements in array
    // (note using dynamic array, which starts at index 0)
    count := 0;
    while count < totalNonlinFlexJoints do begin
      // Read the element number
      ReadLn(o, temp);
      FirstWord(temp, dumStr);

      // Store nonlinear flex joint element in array
      nonlinFJElem[count] := StrToInt(dumStr);

      Inc(count);

      // Check number of nonlinear flex joints in file does not exceed storage limit
      if count > MAX_SPRINGS then begin
        WriteLn(' Error: There are ', count, ' nonlinear flex joints in the model - this is too many for ', PROG_NAME);
        Halt(1);
      end;
    end;

    // Set size of arrays
    SetLength(nonlinFJElem, totalNonlinFlexJoints+1);
    SetLength(nonlinFJTableMom, totalNonlinFlexJoints+1);
    SetLength(nonlinFJTableAng, totalNonlinFlexJoints+1);

    // For each nonlinear flex joint store the moment and angle in tables in
    // arrays based on table number
    i := 0;
    while i < totalNonlinFlexJoints do begin
      dumStr := FindLine(o, 'Table no.:       ' + IntToStr(i+1));
      for loopCounter := 1 to 5 do ReadLn(o, temp);
      j := 0;

      while (temp <> '') and (not EOF(o)) and (j < MAX_SPRINGS) do begin
        // Extract and store Moment
        FirstWord(temp, dumStr);
        nonlinFJTableMom[i,j] := StrToFloat(dumStr);

        // Extract and store Angle
        FirstWord(temp, dumStr);
        nonlinFJTableAng[i,j] := StrToFloat(dumStr);

        ReadLn(o, temp);
        Inc(j);
      end;

      // Trim size of arrays based on number of values in table
      SetLength(nonlinFJTableMom[i], j);
      SetLength(nonlinFJTableAng[i], j);

      Inc(i);
    end;

  end;

  totalFlexJoints := totalLinFlexJoints + totalNonlinFlexJoints;

  // Combine in order, linear and nonlinear flex joint elements into new array
  // flexJoint of size: [totalFlexJoints,0:2], where column:
  //  0: flex joint element
  //  1: 0 = linear, 1 = nonlinear
  //  2: nonlinear table number - 1
  if totalFlexJoints > 0 then begin
    j := 0;
    k := 0;

    for i := 0 to totalFlexJoints - 1 do begin
      if (linFJElem[j] < nonlinFJElem[k]) and (linFJElem[j] <> 0) or (nonlinFJElem[k] = 0) then
        begin
          flexJoint[i,0] := linFJElem[j];
          Inc(j);
        end
      else if (linFJElem[j] > nonlinFJElem[k]) and (nonlinFJElem[j] <> 0) or (linFJElem[j] = 0) then
        begin
          flexJoint[i,0] := nonlinFJElem[k];
          flexJoint[i,1] := 1;
          flexJoint[i,2] := k;
          Inc(k);
        end;
    end;

    // Trim array size
    SetLength(flexJoint, totalFlexJoints, 3);
  end;

  //-----------------
  // Spring elements
  //-----------------
  Reset(o);
  // Find location
  dumStr := FindLine(o, 'No. of Spring Elements');

  for i := 1 to 6 do FirstWord(dumStr, temp);
  Val(temp, totalSprings, code);

  if debug = True then WriteLn (' Springs: ', temp);

  dumStr := FindLine(o, '*** ELEMENT PROPERTIES ***');
  for i := 1 to 5 do ReadLn(o);

  count := 1;
  count2 := 0;

  for i := 1 to totalElements do begin
    ReadLn(o, temp);
    FirstWord(temp, dumStr);
    // Extract element number
    Val(dumStr, elementNo, code);
    if elementNo = 0 then       { To account for the additional blank line in Flexcom version 8.4.4 }
      begin;
      ReadLn(o,temp);
      FirstWord(temp, dumstr);
      Val(dumstr,elementNo,code);
      end;
    FirstWord(temp, dumStr);
    FirstWord(temp, dumStr);
    // Extract potential "Linear Spring" identifier
    dumStr := LowerCase(dumStr);
    FirstWord(temp, dumStr1);
    dumStr1 := LowerCase(dumStr1);

    // Check for linear springs
    if (dumStr = 'linear') and (dumStr1 = 'spring') then
      begin
        spring[count] := elementNo;
        springdata[count,1] := 0;

        for j := 1 to 5 do FirstWord(temp, dumStr);

        Val(dumStr, stiffness, code);
        if code <> 0 then begin; writeln('Error occur in read spring data'); halt(1);end;
        Springdata[count,2] := stiffness;

        count := count + 1;
      end
    // Check for nonlinear springs
    else if (dumStr = 'nonlinear') and (dumStr1 = 'spring') then
      begin
        spring[count] := elementNo;
        springdata[count,1] := 1;

        for j := 1 to 4 do FirstWord(temp, dumStr);
        for j := 1 to 3 do temp[j]:= ' ';

        FirstWord(temp, dumStr);
        Val(dumStr, matlNo, code);

        if matlNo = 0 then begin
          if warnSpring = False then begin
            WriteLn(' Error 100 or greater nonlinear spring material properties');
            WriteLn(' Springs might not associate correctly');
            Inc(warningsCount);
            warningsArr[warningsCount] := ' Error 100 or greater nonlinear spring material properties';
            Inc(warningsCount);
            warningsArr[warningsCount] := ' Springs might not associate correctly ';
            warnSpring := True;
          end;

          count2 := count2 + 1;
          matlNo := 99 + count2;
        end;
        springdata[count,2] := matlNo;

        count := count + 1;
      end;

    // Check number of springs is within limit
    if count > MAX_SPRINGS then begin
      WriteLn ('*** Error, there are ', count, ' springs in the model - this is too many for ', PROG_NAME);
      Halt(1);
    end;

  end;

end;  // GetModelDetails


procedure WriteElem;
{ **************************************************************************** }
// Writes element info to table files of displacement
{ **************************************************************************** }
var
  i, tempInt, node : Integer;
  x, y, z : Real;

begin
  // Set up displacement table
  // Find element number location
  dumStr := FindLine(o, '*** NODAL DATA ***');

  for tempInt := 1 to 4 do
    ReadLn(o);

  for i := 1 to totalNodes do
  begin
    ReadLn(o, node, x, y, z);
    if node = 0 then ReadLn(o, node, x, y, z);
    nodeinfoI[i] := node;
    nodeinfo[i,1] := x;
    nodeinfo[i,2] := y;
    nodeinfo[i,3] := z;
  end;
end; // WriteElem


procedure FactorsInput;
{ **************************************************************************** }
// Input factors.
{ **************************************************************************** }
var
  facType : Char;
  facType1, facType2 : String;
  dumint,NoIntFluidSetOver : Integer;

  { ************************************************************************** }
  // Internal procedure to FactorsInput: writes individual factor columns
  // Read keywords, input, length of number output, number decimal point, unit
  // scale factor.
  { ************************************************************************** }
  procedure DoWrite(facStr : STRING; fnum : INTEGER; units, default : REAL; optional : boolean);

  var
    i, j, k, l , ag, elem, elem1, elem2, number, code, code1, code2 : Integer;
    factor : Array[1..7] of Real;
    tempNum1, tempNum2 : Real;
    g, inputF : Text;
    fileEmpty : bool;
    faclength : integer;

    // Temporary variables
    temp1, temp2, count,dumlen,labelindex : Integer;
    dumDum, dumDum1, tempVar : String;
    dumdum2, stripleft, stripagain : string;

    // key: this is to flag if there is a column of 'I' values following 'G' or vice versa
    key : Boolean;
    // nomidflag: this is to flag if nominal ids are included in INF file
    nomidflag : Boolean;

  begin

    dumStr := FindLine(h, facStr);

    key := False;
    nomidflag := False;

    if (dumStr = NOT_FOUND) and (optional = False) then begin
      WriteLn(' ERROR: INF file missing the keyword ', facStr);
      Halt(1);
    end;

    if (dumStr = NOT_FOUND) and (facStr = 'FORCENODE') and (STT = True) then begin
      WriteLn(' ERROR: Keyword FORCENODE should be used with STT option in INF file');
      WriteLn(' and the local node numbers should be consistent with those in GRD file');
      Halt(1);
    end;

    for k := 1 to 7 do factor[k] := default;

    // Set up defaults
    for i := 1 to totalElements do
      for k := 1 to 7 do factors[i,k] := default;

    facType := ' ';
    if dumStr <> NOT_FOUND then begin
      // Do keyword reading
      ReadLn(h, dumDum);

      // Define yield strength temperature correction curve
      if facStr = 'THERMAL TABLE' then
      begin
          derateFactor := true;
          FirstWord(dumDum, tempVar);
          Val(tempVar, Num_Temperature, code);
          if code <> 0 then begin; writeln('Error in THERMAL TABLE specification'); halt(1); end;
          if Num_Temperature > 10 then begin; writeln('Number of Thermal Table greater than allowed'); halt(1); end;
          for i := 1 to Num_Temperature do
          begin
              ReadLn(h, dumDum);
              FirstWord(dumDum, tempVar);
              Val(tempVar, YieldStrTempCor[i,0,1], code);
              if code <> 0 then begin; writeln('Error in THERMAL TABLE specification'); halt(1); end;
              if YieldStrTempCor[i,0,1] > 20 then begin; writeln('Number of Temperature in THERMAL TABLE greater than allowed'); halt(1); end;
              for j := 1 to round(YieldStrTempCor[i,0,1]) do
              begin
                  ReadLn(h, dumDum);
                  FirstWord(dumDum, tempVar);
                  Val(tempVar, YieldStrTempCor[i,j,1], code);
                  if code <> 0 then
                  begin
                      WriteLn(' ERROR: The number of rows are not matching with the input of temperature curve ', i, '.');
                      Halt(1);
                  end;
                  FirstWord(dumDum, tempVar);
                  Val(tempVar, YieldStrTempCor[i,j,2], code);
                  if code <> 0 then
                  begin
                      WriteLn(' ERROR: The number of rows are not matching with the input of temperature curve ', i, '.');
                      Halt(1);
                  end;
              end;
              ReadLn(h, dumDum);
              FirstWord(dumDum, tempVar);
              if AnsiUpperCase(tempVar) <> 'END' then
              begin
                  WriteLn(' ERROR: The number of rows are not matching with the input of temperature curve ', i, '.');
                  Halt(1);
              end;
          end;//end for loop
      end


      else if facStr = 'LIMIT STATE' then
      begin
          FirstWord(dumDum, tempVar);
          limitState := tempVar;
          if limitState = 'SLS' then
          begin
          factor[2] := 1;
          factor[3] := 0.96;
          factor[5] := 1;
          factor[6] := 1;
          factor[7] := 1;
          end
          else if limitState = 'ULS' then
          begin
          factor[2] := 1.15;
          factor[3] := 0.96;
          factor[5] := 1.1;
          factor[6] := 1.3;
          factor[7] := 1;
          end
          else if limitState = 'ALS' then
          begin
          factor[2] := 1.15;
          factor[3] := 0.96;
          factor[5] := 1;
          factor[6] := 1;
          factor[7] := 1;
          end
          else
          begin
              WriteLn(' ERROR: Must SLS, ULS or ALS for LIMIT STATE in INF file.');
              Halt(1);
          end;
      end

      // Optional keyword. To enable user to replace internal fluid or external in .out file with another internal fluid for multistring modeling.
      else if (facStr = 'INTERNAL FLUID') then
        begin
          IntFluidIndex := True; DumInt := 1;
          //
          FirstWord(dumDum, Factype1);
          Val(FacType1,Factor[1],code);
          if code <> 0 then
           begin;
           IntFluidElementOverride := true;
           faclength := length(FacType1);
           if (facType1 <> '') and (faclength =1) then
            facType := facType1[1]
           else
            facType := ' ';
          end
          else
          begin;
          for k := 2 to 4 do begin
            FirstWord(dumDum, facType1);
            Val(facType1, factor[k], code);
          end;
           for k := 1 to 4 do
            factors[dumint,k] := factor[k];
          repeat;
          Readln(h,dumdum);
          inc(dumint);
          Firstword(dumdum,facType1);
          Val(FacType1,factor[1],code);
          if code <> 0 then facType := ' '
          else
          begin;
          for k := 2 to 4 do begin
            FirstWord(dumDum, facType1);
            Val(facType1, factor[k], code);
          end;
           for k := 1 to 4 do
            factors[dumint,k] := factor[k];
          end;
          until((EOF(h) = true) or (code <> 0));
          NoIntFluidSetOver := dumint;
          end;
        end

      else if (facStr = 'EXTERNAL FLUID') then
        begin;
        ExtFluidIndex := True;
        FirstWord(dumDum, Factype1);
        Val(FacType1,Factor[1],code);
         if code <> 0 then
           begin;
           Faclength := length(FacType1);
           if (facType1 <> '') and (Faclength =1) then
            facType := facType1[1]
           else
            facType := ' ';
           end
          else
          begin;
            for k := 2 to 4 do
            begin
            FirstWord(dumDum, facType1);
            Val(facType1, factor[k], code);
            end;
          for i := 1 to totalelements do
           for k := 1 to 4 do
            factors[i,k] := factor[k];
          end;
        end

      else
        begin
          FirstWord(dumDum, facType1);
          faclength := length(FacType1);
          // This statement safeguards against 2HSTRE3D crashing in case there are spaces between lines in INF File (AG, 27/1/2009)
          if (facType1 <> '') and (faclength = 1)  then
            // Picks up first character from the whole line
            facType := facType1[1]
          else
            begin;
            if facStr <> 'STT' then
             begin;
             writeln('Must have a valid D,G,I,L specified for ',facStr);
             halt(1);
             facType := ' ';
             end;
            end;

          if facStr = 'NOMINAL ID' then nomidflag := True;
          if facStr = 'COATING OD' then coatodflag := True;
        end;

      repeat

        case facType of

          // D is default value
          'D' : begin
            if facStr = 'SAFETY CLASS' then
            begin
                ReadLn(h, dumDum);
                FirstWord(dumDum, tempVar);
                tempVar := UpperCase(tempVar);
                if tempVar = 'HIGH' then
                    factor[1] := 1.26
                else if tempVar = 'NORMAL' then
                    factor[1] := 1.14
                else if tempVar = 'LOW' then
                    factor[1] := 1.04
                else
                begin
                    WriteLn('Wrong input for SAFETY CLASS!');
                    Halt(1);
                end;
            end
            else
            begin
                try
                  for k := 1 to fnum do Read(h, factor[k]);
                except on E: Exception do begin;
                  writeln('Error in reading default line ', facStr, ' Error message ', E.Message );
                  halt(1);
                end;
                end;
                // Default values
                ReadLn(h);
            end;

            if debug then WriteLn(' Reading default of , ', factor[1], ' for the ', facStr);

            for i := 1 to totalElements do begin
              for k := 1 to fnum do factors[i,k] := factor[k];
            end
          end;

          // I is individual value
          'I' : begin
            FirstWord(dumDum, facType2);
            // Number of individual values
            Val(facType2, number, code);

            if code <> 0 then begin
              WriteLn(' Syntax error in INF file, individual command block line for ', dumStr);
              Halt(1);
            end;

            ReadLn(h, dumDum);
            FirstWord(dumDum, tempVar);
            // Area that need to be modified for *LABELS read in...
            Val(tempVar, elem, code);
            count := 0;
            key := False;

            while (code = 0) and (count < 100000) do begin

              if facStr = 'SAFETY CLASS' then
              begin
                  FirstWord(dumDum, tempVar);
                  tempVar := UpperCase(tempVar);
                  if tempVar = 'HIGH' then
                      factor[1] := 1.26
                  else if tempVar = 'NORMAL' then
                      factor[1] := 1.14
                  else if tempVar = 'LOW' then
                      factor[1] := 1.04
                  else
                  begin
                      WriteLn('Wrong input for SAFETY CLASS!');
                      Halt(1);
                  end;
              end
              else
              begin
                  for k := 1 to fnum do begin
                      FirstWord(dumDum, tempVar);
                      Val(tempVar, factor[k], code);
                  end;
              end;

              for i := 1 to totalElements do
                for k := 1 to fnum do
                  if elemInfo[i,1] = elem then factors[i,k] := factor[k];

              count := count + 1;

              ReadLn(h, dumDum);
              dumDum1 := dumDum;
              FirstWord(dumDum, tempVar);

              if (tempVar = 'G') or (tempVar = 'L') then begin
                facType := tempVar[1];
                dumDum := dumDum1;
                key := True;
                break;
              end;

              Val(tempVar, elem, code);

            end;

            if count <> number then begin
              WriteLn('Error: number of individual (I) values incorrectly specified for ', facStr, ' in INF');
              Halt(1);
            end;

            if (tempVar <> 'G') and (tempVar <> 'L') then break;

          end;

          // G is group of values
          'G' : begin
            FirstWord(dumDum, facType2);
            // Number of group values
            Val(facType2, number, code);

            if code <> 0 then begin
              WriteLn(' Syntax error in INF file, group command block line for ', dumStr);
              Halt(1);
            end;

            count := 0;
            key := False;

            While (key = false) and (not EOF(h)) and (count < 100000) do
            begin;

            ReadLn(h, dumDum);
            dumDum2 := dumDum;
            FirstWord(dumDum, tempVar);
            if (tempVar = 'I') or (tempVar = 'L') then
               begin;
               facType := tempVar[1];
               dumDum := dumDum2;
               key := True;
               break;
               end;
            Val(tempVar, elem1, code1);
            if (code1 <> 0) then
              begin;
              if count > number then writeln(' Error in read ', facStr, ' - Number Expected');
              key := true;
              end
              else
              begin;
              FirstWord(dumDum, tempVar);
              Val(tempVar, elem2, code2);
              if code2 <> 0 then
                begin;
                writeln(' Error in read ',facStr, ' - Number Expected');
                halt(1);
                key := true;
                end;
              end;

            if key = true then break;

            if elem2 < elem1 then
              begin;
              WriteLn(' Error in read ', facStr, ' - Group Elements Must Increase');
              Halt(1);
              end;

            if facStr = 'SAFETY CLASS' then
              begin
                  FirstWord(dumDum, tempVar);
                  tempVar := UpperCase(tempVar);
                  if tempVar = 'HIGH' then factor[1] := 1.26
                  else if tempVar = 'NORMAL' then factor[1] := 1.14
                  else if tempVar = 'LOW' then factor[1] := 1.04
                  else
                  begin
                      WriteLn('Wrong input for SAFETY CLASS!');
                      Halt(1);
                  end;
              end
            else
              begin
                  for k := 1 to fnum do begin
                      FirstWord(dumDum, tempVar);
                      Val(tempVar, factor[k], code);
                  end;
              end;

              for i := 1  to totalElements do begin
                if (elemInfo[i,1] >= elem1) and (elemInfo[i,1] <= elem2) then
                  for k:=1 to fnum do factors[i,k] := factor[k];
              end;

              Inc(count);

            end;

            if count <> number then begin
              WriteLn(' Error: number of group (G) values incorrectly specified for ', facStr, ' in INF');
              Halt(1);
            end;

            if (tempVar <> 'I') and (tempVar <> 'L') then break;

          end;

                     // G is group of values
          'L' : begin
            FirstWord(dumDum, facType2);
            // Number of group values
            Val(facType2, number, code);

            if code <> 0 then begin
              WriteLn(' Syntax error in INF file, label command block line for ', dumStr);
              Halt(1);
            end;

            count := 0;
            key := False;

            While (key = false) and (not EOF(h)) and (count < 100000) do
            begin;

            ReadLn(h, dumDum);
            dumDum2 := dumDum;
            FirstWord(dumDum, tempVar);
            if (tempVar = 'I') or (tempVar = 'G') then begin
               facType := tempVar[1];
               dumDum := dumDum2;
               key := True;
               break;
            end;
              if pos('}',dumdum2) <> 0 then
                begin;
                dumdum := rightstr(dumdum2, Length(dumdum2) - pos('}',dumDum2));
                stripleft := leftstr(dumdum2, pos('}',dumdum2)-1)
                end
              else begin; key := true; stripleft := 'Error';  end;
              if pos('{',stripleft) <> 0 then
                begin;
                dumlen := length(stripleft);
                stripagain := rightstr(stripleft,dumlen - pos('{',stripleft))
                end
              else key := true;

              labelindex := 0;
              if key = false then
              begin;
              for l := 1 to TotalLabelE do
                if Uppercase(stripagain) = Uppercase(LabelElem[l]) then  labelindex := l;
              if labelindex = 0 then
                    begin;
                    writeln('WARNING LABEL: ',stripagain,' UNKNOWN');
                    Inc(warningsCount);
                    warningsArr[warningsCount] := 'WARNING LABEL: ' + stripagain + ' UNKNOWN';
                    elem1 := -1;
                    elem2 := 0;
                    end
              else
                begin;
                elem1 := LabelElemI[labelindex,1];
                elem2 := LabelElemI[labelindex,2];
                end;
              end;

            if key = true then break;

            if elem2 < elem1 then begin
              WriteLn(' Error in read ', facStr, ' - Label Elements Must Increase');
              Halt(1);
            end;

            if facStr = 'SAFETY CLASS' then
              begin
                  FirstWord(dumDum, tempVar);
                  tempVar := UpperCase(tempVar);
                  if tempVar = 'HIGH' then factor[1] := 1.26
                  else if tempVar = 'NORMAL' then factor[1] := 1.14
                  else if tempVar = 'LOW' then factor[1] := 1.04
                  else
                  begin
                      WriteLn('Wrong input for SAFETY CLASS!');
                      Halt(1);
                  end;
            end
            else
              begin
                  for k := 1 to fnum do begin
                      FirstWord(dumDum, tempVar);
                      Val(tempVar, factor[k], code);
                  end;
              end;

              for i := 1  to totalElements do begin
                if (elemInfo[i,1] >= elem1) and (elemInfo[i,1] <= elem2) then
                  for k:=1 to fnum do factors[i,k] := factor[k];
              end;

              Inc(count);

            end;

            if count <> number then begin
              WriteLn(' Error: number of label (L) values incorrectly specified for ', facStr, ' in INF');
              Halt(1);
            end;

            if (tempVar <> 'I') and (tempVar <> 'G') then break;

          end;

          else
             begin;
             if Factype = ' ' then continue
             else
              begin;
              writeln('Factor Type :', Factype ,' is unknown for ',FacStr);
              halt(1);
              end;
             end;

        end;  // End of case statement

        // If dumDum has already been read while reading through the I or G column,
        // then next line should not be read to obtain dumDum.
        if key = False then ReadLn(h, dumDum);

        FirstWord(dumDum, facType1);
        faclength := length(facType1);
        if (facType1 <> '') and (faclength=1) then
          facType := facType1[1]
        else
          facType := ' ';

      until not ((facType = 'I') or (facType = 'G') or (facType = 'D') or (facType = 'L'));

    end;

    // Leave files alone
    if optional = False then
      begin

        // Into array structure
        for i := 1 to totalElements do begin
          if facStr = 'PRES ADJ'then Pinfo[i,1] := factors[i,1]/units;
          if facStr = 'TENSION' then pipeInfo[i,1] := factors[i,1]/units;
          if facStr = 'BENDING'then pipeInfo[i,2] := factors[i,1]/units;
          if facStr = 'HOOP' then pipeInfo[i,3] := factors[i,1]/units;
          if facStr = 'YIELD' then pipeInfo[i,4] := factors[i,1]/units;
          if facStr = 'NOMINAL OD' then pipeInfo[i,5] := factors[i,1]/units;
          if facStr = 'CORR INT' then pipeInfo[i,6] := factors[i,1]/units;
          if facStr = 'CORR EXT' then pipeInfo[i,7] := factors[i,1]/units;
          if facStr = 'TENS ADJ' then pipeInfo[i,8] := factors[i,1]/units;
          if facStr = 'ULTI STR' then pipeInfo[i,10] := factors[i,1]/units;
          if facStr = 'YOUNGS' then pipeInfo[i,11] := factors[i,1]/units;
          if facStr = 'OVALITY' then pipeInfo[i,12] := factors[i,1]/units;
		  if facStr = 'FAB PARAM' then
          begin
              BurstK[i] := factors[i,1]/units;
              Alphafab[i] := factors[i,2]/units;
              DNVFactors[i,4] := factors[i,2]/units;
              FcM4[i] := factors[i,3]/units;
              pipeInfo[i,12] := factors[i,4]/units; //Ovality
              DNVMatSpec[i,6] := factors[i,4]/units*2; //Ovality for DNV
          end;
          if facStr = 'MAT SPEC' then
          begin
              pipeInfo[i,4] := factors[i,1]/1000000;    // Yield strength  (MPa)
              DNVMatSpec[i,1] := factors[i,2]/1000000;  //Tensile strength for DNV   (MPa)
              pipeInfo[i,10] := factors[i,2]/1000000;   //Tensile strength (MPa)
              DNVMatSpec[i,2] := factors[i,3]/1000000;  //Emod for DNV (MPa)
              pipeInfo[i,11] := factors[i,3]/1000000;   //Emod (MPa)
              pipeInfo[i,13] := factors[i,4]/units;     //Poisson ratio
              DNVMatSpec[i,3] := factors[i,4]/units;    //Poisson ratio for DNV
              if globalFactor then
              begin
                  DNVMatSpec[i,4] := (1-temperatureFactor[i])*pipeInfo[i,4];
                  DNVMatSpec[i,5] := (1-temperatureFactorTensile[i])* DNVMatSpec[i,1];
              end
              else
              begin
                  GetYieldStrFac(i);
                  DNVMatSpec[i,4] := (1-YieldStr_Fac)*pipeInfo[i,4];
                  DNVMatSpec[i,5] := (1-YieldStr_Fac)* DNVMatSpec[i,1];
              end;
              if ((factors[i,1] < 1000000) Or (factors[i,2] < 1000000) Or (factors[i,3] < 1000000)) then
              begin
                  WriteLn('Please check your input for "MAT SPEC", unit should be Pa!');
                  halt(1);
              end;
          end;
          if facStr = 'COLLAPSE EQ' then
          begin
              if factors[i,1] = 1 then
                  Collapse_Type[i] := true
              else if factors[i,1] = 0 then
                  Collapse_Type[i] := false
              else
              begin
                  WriteLn(' Error: COLLAPSE Equation type must be 0 or 1 in INF');
                  Halt(1);
              end;
          end;

          if facStr = 'OVALITY' then
          begin
              pipeInfo[i,12] := factors[i,1]/units;
              DNVMatSpec[i,6] := factors[i,1]/units*2; // factor of 2 as ...
          end;
          if facStr = 'POISSON' then pipeInfo[i,13] := factors[i,1]/units;
          if facStr = 'SAFETY CLASS' then DNVFactors[i,1] := factors[i,1]/units;
          if facStr = 'LIMIT STATE' then
          begin
              DNVFactors[i,2] := factor[2];
              DNVFactors[i,3] := factor[3];
              DNVFactors[i,5] := factor[5];
              DNVFactors[i,6] := factor[6];
              DNVFactors[i,7] := factor[7];
          end;
          if facStr = 'TEMPERATURE VALUE' then
          begin
              temperature[i,1] := factors[i,1];
              temperature[i,2] := factors[i,2];
          end;
        end;
    end
    // Optional is true
    else
      begin

        if facStr = 'ZERO PRESSURE CHECK' then
          for i := 1 to totalElements do zeroPressureCheck[i] := Round(factors[i,1]);

        if facStr = 'STRAINAF' then
          for i := 1 to totalElements do SAFactor[i] := factors[i,1]/units;
        // For multistring modeling, static tension without offset, waves and current is needed
        if facStr = 'NOMINAL STATIC TENSION' then
          for i := 1 to totalElements do
            StatTen[i] := factors[i,1]*units;

        // Overwrites data picked up from .out file in procedure 'geometry' into variable 'pipeInfo'
        if (facStr = 'NOMINAL ID') then
          if (nomidflag = True) then
          begin;
          for i := 1 to totalElements do
            pipeInfo[i,9] := factors[i,1];
          end;

        // Coating OD data
        if (facStr = 'COATING OD') and (coatodflag = true) then
          begin;
          for i := 1 to totalElements do
           pipeInfo[i,14] := factors[i,1]/units;
          end;

        if (facStr = 'INTERNAL FLUID') and (IntFluidIndex = true) then
          begin
          if IntFluidElementOverride = false then
           begin;
           for i := 1 to NoIntFLuidSetover do
             begin;
             FluidSetNo := round(factors[i,1]);
             if FluidSetNo > MAX_FLUIDS then begin; writeln('Number of Fluids in Flexcom file greater than allow in 2HSTRE3D'); halt(1); end;
             IntFluid[FluidSetNo,1] := factors[i,2];
             IntFluid[FluidSetNo,2] := factors[i,3];
             IntFluid[FluidSetNo,3] := factors[i,4];
             end;
           end;
          if IntFluidElementOverride = true then
           begin;
           for i := 1 to totalElements do
            begin;
            IntFluidElement[i,1] := factors[i,1]; // Elevation
            IntFluidElement[i,2] := factors[i,2]; // Fluid Density
            IntFluidElement[i,3] := factors[i,3]; // Pressure at top.
            end;
           end;
        end;

        if (facStr = 'EXTERNAL FLUID') and (ExtFluidIndex = True) then
        begin
          for i := 1 to totalElements do
          begin;
          ExtFluid[i,1] := factors[i,1];
          ExtFluid[i,2] := factors[i,2];
          ExtFluid[i,3] := factors[i,3];
          end;
        end;

        if facStr = 'VMTT' then begin
          if (dumstr <> NOT_FOUND) then
           begin;
            Writeln('ERROR: VMTT keyword removed from 2HSTRE3D');
            halt(1);
           end;
         end;

        if facStr = 'OSTT' then
          begin;
          for i := 1 to totalElements do OSTTElem[i] := round(factors[i,1]);
          if (dumStr <> NOT_FOUND) then OSTT := true;
          if (OSTT = true) and (STT = false) then
           begin; writeln('ERROR: STT option need to get output timetraces');halt(1); end;
          end;

        if facStr = 'MULTS DNV' then begin
          for i := 1 to totalElements do
            for k := 1 to 2 do
              DNVMults[i,k] := factors[i,k];
        end;

        if facStr ='FORCENODE' then begin
          if debug then WriteLn('Force nodes elemSwitch activation');
          for i := 1 to totalElements do elemSwitch[i] := factors[i,1];
        end;

        if facStr ='FACTORS DNV' then begin
          if debug then WriteLn('DNV factors read in ');
          for i := 1 to totalElements do
            for k := 1 to 7 do
              DNVFactors[i,k] := factors[i,k];
        end;

        if facStr = 'MAT SPEC DNV' then begin
          if debug then WriteLn('DNV Materials Read In');
          for i := 1 to totalElements do
          begin
              DNVMatSpec[i,1] := factors[i,1]/1000000;    //convert to Mpa
              DNVMatSpec[i,2] := factors[i,2]/1000000;    //convert to Mpa
              DNVMatSpec[i,3] := factors[i,3];
              DNVMatSpec[i,4] := factors[i,4]/1000000;    //convert to Mpa
              DNVMatSpec[i,5] := factors[i,5]/1000000;    //convert to Mpa
              DNVMatSpec[i,6] := factors[i,6];
              if ((factors[i,1] < 1000000) Or (factors[i,2] < 1000000)) And (DNVCodeCheck) then
              begin
                  WriteLn('Please check your input for "MAT SPEC DNV", unit should be Pa!');
                  halt(1);
              end;
          end;
        end;

        if facStr = 'TEMPERATURE FACTOR' then
        begin
            if dumStr <> NOT_FOUND then globalFactor := true;
            for i := 1 to totalElements do
            begin
	       temperatureFactor[i] := factors[i,1];
                temperatureFactorTensile[i] := factors[i,2];
            end;
        end;
		
        if facStr = 'STT' then begin
          for i := 1 to totalElements do STTElem[i] := factors[i,1];
          if dumStr <> NOT_FOUND then STT := True;
        end;

        if facStr = 'PRES DNV' then begin
          if dumStr <> NOT_FOUND then DNVPinput := True;
            for i := 1 to totalElements do
            begin
              DNVpres[i,1] := factors[i,1]/1000000;    //convert to Mpa
              DNVpres[i,2] := factors[i,2];
              DNVpres[i,3] := factors[i,3];
              DNVpres[i,4] := factors[i,4]/1000000;    //convert to Mpa
              DNVpres[i,5] := factors[i,5];
              DNVpres[i,6] := factors[i,6];
            end;
        end;

        if facStr = 'FLAG' then
          for i := 1 to totalElements do flag := Round(factors[i,1]);

      end;

  end;  // DoWrite

begin

  // Arguments for DoWrite: name, length, dp, fnum, units, default, optional

  DoWrite('INTERNAL FLUID',3,1,0,True); // Internal fluid data
  DoWrite('EXTERNAL FLUID',3,1,0,True); // Internal fluid data
  DoWrite('ZERO PRESSURE CHECK',1,1,0,True);  // Flag for zero top pressure check on von Mises stress. Default is 0-NotCheckingZeroTopPressure
  DoWrite('PRES ADJ',1,1e6,0,False);    // Pressure adjustment
  DoWrite('TENSION',1,1,1,False);       // Tension
  DoWrite('BENDING',1,1,1,False);       // Bending
  DoWrite('HOOP',1,1,1,False);          // Hoop
  if not APISTD2RDCodeCheck then            // disable "YIELD" for API-STD-2RD
    DoWrite('YIELD',1,1e6,0,False);     // Yield
  if APISTD2RDCodeCheck then begin
    DoWrite('SAFETY CLASS',1,1,0,False);// SAFETY CLASS
    globalFactor := false;
    derateFactor := false;
    DoWrite('TEMPERATURE FACTOR',2,1,0,true);// TEMPERATURE FACTOR for yield/tensile strength de-rating
    if globalFactor = false then
    begin
        DoWrite('THERMAL TABLE',1,1,0,False);     // Yield Strength Temperature Correction
        DoWrite('TEMPERATURE VALUE',2,1,0,False); // Temperature
    end;
    if (globalFactor = false) and (derateFactor = false) then
    begin
        WriteLn('Please define de-rating factor for Yield and Tensile Strength!');
        halt(1);
    end;
    DoWrite('COLLAPSE EQ',1,1,0,False);  // Collapse Equation option (Eq.2 or Eq.5 in API-STD-2RD)
    DoWrite('MAT SPEC',4,1,0,False);    // Define yield strength, tensile strength, E, poisson ratio
    DoWrite('LIMIT STATE',1,1,0,False);  // Limit state SLS, ULS or ALS
    DoWrite('STRAINAF',1,1,-1,True);      // Define strain amplification factor, Optional Only for Method 4.
    DoWrite('FAB PARAM',4,1,0,False);    // Define Mechanical variability k, Alpha Fab, collapse factor fc, ovality
  end;
  DoWrite('NOMINAL OD',1,1,0,False);    // Nominal pipe outer diameter
  DoWrite('NOMINAL ID',1,1,0,True);     // Nominal pipe inner diameter
  DoWrite('COATING OD',1,1,0,True);     // Coating pipe outer diameter
  DoWrite('CORR INT',1,1e-3,0,False);   // Internal corrosion
  DoWrite('CORR EXT',1,1e-3,0,False);   // External corrosion
  DoWrite('TENS ADJ',1,1,0,False);      // Tension adjustment
  DoWrite('NOMINAL STATIC TENSION',1,1e3,0,True); // Static tension without waves, current or offset
  if not APISTD2RDCodeCheck then            // 1.52 disable "FACTORS DNV" and "MAT SPEC DNV" for API-STD-2RD
  begin
    DoWrite('FACTORS DNV',7,1,1,True);    // DNV Factors
    DoWrite('MAT SPEC DNV',6,1,0,True);   // DNV Material Specification
  end;
  DoWrite('MULTS DNV',2,1,1,True);      // DNV Multipliers
  DoWrite('STT',1,1,1,True);            // SST Selection
  DoWrite('VMTT',1,1,0,True);           // Check if old keyword VMTT is in file
  DoWrite('OSTT',1,1,0,True);           // Output STT timetrace option.
  DoWrite('FORCENODE',1,1,0,True);      // Forcenode to 1, 2 - zero no forcing
  DoWrite('PRES DNV',6,1,0,True);       // DNV design pressure information
  DoWrite('FLAG',1,1,0,True);           // Flag input

  if ISOCodeCheck then begin
    DoWrite('ULTI STR',1,1e6,0,False);  //  Ultimate Tensile Strength
    DoWrite('YOUNGS',1,1e6,0,False);    //  Youngs Modulus
    DoWrite('OVALITY',1,1,0,False);     //  Initial Ovality
    DoWrite('POISSON',1,1,0,False);     //  Poisson's Ratio
  end;
  dumstr := FindLine(h, 'SSCURVE');
  if (dumStr <> NOT_FOUND) then         //  Warn that Old key still present
    begin;
    Inc(warningsCount);
    warningsArr[warningsCount] := WARN_SSCURVE;
    end;

  CloseFile(h);

end;  // FactorsInput


{ **************************************************************************** }
// Calculates node height.
{ **************************************************************************** }
procedure NodeHeight(elem : Integer; var startElev, endElev : Real; elevtype : integer);

{ elevtype, 1 - average of min, max, 2 - static position, 3 - min X, 4 - max X }

var
  startFlag, endFlag, tempNode, node1, node2,j : Integer;
  elev : Real;

begin
  node1 := elemInfo[elem,2];
  node2 := elemInfo[elem,3];
  startFlag := 0;
  endFlag := 0;
  j := 0;

  repeat
    inc(j,1);
    tempNode := nodeinfoI[j];
    if elevtype = 1 then
      elev := (nodeinfo[j,4]+ nodeinfo[j,5])/2;
    if elevtype = 2 then
      elev := nodeinfo[j,1];
    if elevtype = 3 then
      elev := nodeinfo[j,4];
    if elevtype = 4 then
      elev := nodeinfo[j,5];

    if tempNode = node1 then begin
      startElev := elev;
      startFlag := 1;
    end;

    if tempNode = node2 then begin
      endElev := elev;
      endFlag := 1;
    end;
  until (startFlag = 1) and (endFlag = 1);

end;  // NodeHeight


{ **************************************************************************** }
// Calculates element length.
{ **************************************************************************** }
procedure ElementLength(elem : Integer; var elemLength : Real);

var
  startFlag, endFlag, tempNode, node1, node2,j : Integer;
  stx, sty, stz, enx, eny, enz, x, y, z : Real;

begin
  node1 := elemInfo[elem,2];
  node2 := elemInfo[elem,3];
  startFlag := 0;
  endFlag := 0;
  j := 0;
  stx := 0; sty := 0; stz := 0;
  enx := 0; eny := 0; enz := 0;

  repeat
    inc(j,1);
    if j > totalnodes then begin; writeln('Error in Node Positions'); halt(1); end;
    tempnode := nodeInfoI[j];
    x := nodeinfo[j,1];
    y := nodeInfo[j,2];
    z := nodeinfo[j,3];

    if tempNode = node1 then begin
      stx := x;
      sty := y;
      stz := z;
      startFlag := 1;
    end;

    if tempNode = node2 then begin
      enx := x;
      eny := y;
      enz := z;
      endflag := 1;
    end;
  until (startFlag = 1) and (endFlag = 1);

  elemLength := Sqrt(Sqr(stx - enx) + Sqr(sty - eny) + Sqr(stz - enz));

end;  // ElementLength


{ **************************************************************************** }
// Calculates tension corrections.
{ **************************************************************************** }
procedure TensionCorrection;

type
  T1DReArray = Array[1..2] of Real;

var
  suppnode, suppelem1, node, node1, i : Integer;
  elev, accum, tubmassperlen,
  longness : Real;
  tensioncorr : T1DReArray;

begin
  // Calculate tension corrections
  // Check to see where tubing support elevation is
  suppnode := nodeinfoI[1];
  for i := 1 to totalNodes do
    begin;
     node := nodeinfoI[i];
     elev := nodeinfo[i,1];
     if suppelev <= elev then begin
      suppnode := node;   // Tubing support elevation is at current node
      break; //
    end;
  end;
  
  // Find element attached to suppnode

  i := 0;
  repeat;
    inc(i,1);
    node := eleminfo[i,2];
    node1 := eleminfo[i,3];
  until (node = suppnode) or (node1 = suppnode) or (i >= totalElements);

  accum := -tubzerotens;

  // Element numbering may not necessarily begin at 1. Previously, assumption was element numbering always begins at 1. (AG, 27/1/2009)
  suppelem1 := i;

  // Calculate tension correction for top supported section
  for i := 1 to suppelem1 do begin
    ElementLength(i, longness);
    tubmassperlen := Pipeinfo[i,8];
    tensioncorr[1] := accum;
    tensioncorr[2] := accum - tubmassperlen*longness*GRAVITY;
    accum := tensioncorr[2];
    tensCorr[i] := 0.5*(tensioncorr[1] + tensioncorr[2])/1000;
  end;

  // Calculates for bottom supported section
  accum := 0;

  for i := totalElements downto suppelem1 + 1 do begin
    ElementLength(i,longness);
    tubmassperlen := Pipeinfo[i,8];
    tensioncorr[1] := accum + (tubmassperlen *
    longness * 9.81);
    tensioncorr[2] := accum;
    accum := tensioncorr[1];
    tensCorr[i] := (tensioncorr[1] + tensioncorr[2])/2/1000;
  end;

end; // TensionCorrection


{ **************************************************************************** }
// Reads forces from .tab file.
{ **************************************************************************** }
procedure ReadForces(searchString : String; skipLines : Integer;
                    forceTitle, units : String);

type
  T1DDbleArray = array [1..4] of Double;

var
  i, sprCount, flexJntCount, tempInt, elem : Integer;
  temp1, temp2 : Real;
  force : T1DDbleArray;
  bendFac, factor : Real;
  dumpstr,dumpstr2 : string;

begin
  // Find max/min data location
  dumStr := FindLine(t, searchString);

  if Uppercase(dumstr) = 'NOT FOUND' then
   begin;
   writeln('Error ',SearchString,' is not found in Tab file');
   halt(1);
   end;

  for tempInt := 1 to skipLines do ReadLn(t);

  Lheader[1] := Lheader[1] + forceTitle;
  Lheader[2] := Lheader[2] + '   Min     Max  ';
  Lheader[3] := Lheader[3] + units;
  rRef := rRef + 2;

  sprCount := 1;
  flexJntCount := 0;

  for i := 1 to totalElements do begin

    Readln(t,dumpstr);
    Firstword(dumpstr,dumpstr2);
    Val(dumpstr2,elem,code);
    if code <> 0 then
          begin;
          writeln('Error ',SearchString,' is not correctly input in Tab file');
          halt(1);
          end;
    // Bending factor
    bendFac := pipeInfo[i,2];

    if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) then
      begin
        if elem = spring[sprCount] then Inc(sprCount)
        else if flexJntCount < High(flexJoint) then Inc(flexJntCount)
      end

    else
      begin
        if forcetitle <> '   BEND RADIUS  ' then
          begin
            FirstWord(dumpstr,dumpstr2);
            Val(dumpstr2,force[1],code);
            FirstWord(dumpstr,dumpstr2);
            Val(dumpstr2,force[2],code);
            FirstWord(dumpstr,dumpstr2);
            FirstWord(dumpstr,dumpstr2);
            FirstWord(dumpstr,dumpstr2);
            Val(dumpstr2,force[3],code);
            FirstWord(dumpstr,dumpstr2);
            Val(dumpstr2,force[4],code);

            // Bending moments are multiplied by the bending factor
            if (rRef >= 9) and (rRef < 14) then
              factor := bendFac
            else
              factor := 1.0;

            loads[i,rRef,1,currentFileNo] := force[1]*factor;
            loads[i,rRef+1,1,currentFileNo] := force[2]*factor;
            loads[i,rRef,2,currentFileNo] := force[3]*factor;
            loads[i,rRef+1,2,currentFileNo] := force[4]*factor;
          end

        else
          begin
            FirstWord(dumpstr,dumpstr2);
            Val(dumpstr2,force[1],code);
            FirstWord(dumpstr,dumpstr2);
            Val(dumpstr2,force[2],code);
            FirstWord(dumpstr,dumpstr2);
            FirstWord(dumpstr,dumpstr2);
            FirstWord(dumpstr,dumpstr2);
            Val(dumpstr2,force[3],code);
            FirstWord(dumpstr,dumpstr2);
            Val(dumpstr2,force[4],code);

            { Record curvature directly in the loads array }
            loads[i,rRef+2,1,currentFileNo] := force[1];
            loads[i,rRef+3,1,currentFileNo] := force[2];
            loads[i,rRef+2,2,currentFileNo] := force[3];
            loads[i,rRef+3,2,currentFileNo] := force[4]; 

            for tempInt := 1 to 4 do begin
              if (force[tempInt]=0) or (1/force[tempInt] > 99999.0) then
                force[tempInt] := 1/99999.0;
            end;

            loads[i,rRef+1,1,currentFileNo] := 1/force[1];
            loads[i,rRef,1,currentFileNo] := 1/force[2];
            loads[i,rRef+1,2,currentFileNo] := 1/force[3];
            loads[i,rRef,2,currentFileNo] := 1/force[4];
        end

    end;

  end;

end; // ReadForces


{ **************************************************************************** }
// Reads displacements from .tab file.
{ **************************************************************************** }
procedure ReadDisplacements(searchStr : String; index1,index2: integer);

var
  i, j, tempInt, node, errorChk,tempFrac : Integer;
  disp1, disp2 : Real;
  dump, textDump : String;

begin
  // Find min/max data location
  dumStr := FindLine(t, searchStr);

  if UpperCase(dumstr) = 'NOT FOUND' then
   begin;
   writeln('Error ',Searchstr,' is not found in Tab file');
   halt(1);
   end;
  if searchstr = 'ANGULAR Y DISP ENVELOPE' then
   begin;
   readln(t,textdump);
   readln(t,textdump);
   if POS('DOF 6',textdump) <> 0 then
    begin;
    index1 := 14;
    index2 := 15;
    end;
   end;
  if searchstr = 'ANGULAR Z DISP ENVELOPE' then
   begin;
   readln(t,textdump);
   readln(t,textdump);
   if POS('DOF 5',textdump) <> 0 then
    begin;
    index1 := 12;
    index2 := 13;
    end;
   end;
  // Would you believe it, there is a difference between q-static and dynamic tab files !!!!!
  // It was required to re-write this section to make it more robust.
  // Check to see how many rows to read (tab file contains data for two nodes per row)
  // Find number of double node rows (discount remainder in division)
  tempInt := totalNodes Div 2;
  // Determine whether there is a last row containing just one node
  tempFrac := Ceil(Frac(totalNodes / 2));

  // Read double columns
  i := 1;
  j := 0;

  while (i <= tempInt) and (not EOF(t)) do begin
    ReadLn(t, textDump);
    FirstWord(textDump, dump);

    // Convert first word of line to numeric (if node number, errorChk will be zero)
 		Val(dump, node, errorChk);

    // If line with node data found
    if errorChk = 0 then begin

      FirstWord(textDump, dump);
      disp1 := StrToFloat(dump);

      FirstWord(textDump, dump);
      disp2 := StrToFloat(dump);

      inc(j,1);
      nodeinfo[j,index1] := disp1;
      nodeinfo[j,index2] := disp2;

      FirstWord(textDump, dump);
      node := StrToInt(dump);

      FirstWord(textDump, dump);
      disp1 := StrToFloat(dump);

      FirstWord(textDump, dump);
      disp2 := StrToFloat(dump);

      nodeinfo[j+tempInt+tempFrac,index1] := disp1;
      nodeinfo[j+tempInt+tempFrac,index2] := disp2;

      Inc(i);

    end;

  end;

  // Read single column if any
  if tempFrac <> 0 then begin

    ReadLn(t, textDump);

    FirstWord(textDump, dump);
    node := StrToInt(dump);

    FirstWord(textDump, dump);
    disp1 := StrToFloat(dump);

    FirstWord(textDump, dump);
    disp2 := StrToFloat(dump);

    inc(j,1);
    nodeinfo[j,index1] := disp1;
    nodeinfo[j,index2] := disp2;

  end;

end;  // ReadDisplacements


{ **************************************************************************** }
// Creates displacements table for .tb1 file.
{ **************************************************************************** }
procedure Displacements;

begin

  if (flexVer1 >= 8) and (ExtremeData = False) then
    begin
      ReadDisplacements('VERTICAL X DISP ENVELOPE', 4,5);
      ReadDisplacements('HORZ Y DISP ENVELOPE', 6,7);
      ReadDisplacements('HORZ Z DISP ENVELOPE', 8,9);
      ReadDisplacements('TORSIONAL DISP ENVELOPE', 10,11);
      ReadDisplacements('ANGULAR Y DISP ENVELOPE', 12,13);
      ReadDisplacements('ANGULAR Z DISP ENVELOPE', 14,15);
    end

  else if (FlexVer1 >= 8) and (ExtremeData = True) then
    begin
      ReadDisplacements('VERTICAL X DISP EXTREME', 4,5);
      ReadDisplacements('HORZ Y DISP EXTREME', 6,7);
      ReadDisplacements('HORZ Z DISP EXTREME', 8,9);
      ReadDisplacements('TORSIONAL DISP EXTREME', 10,11);
      ReadDisplacements('ANGULAR Y DISP EXTREME', 12,13);
      ReadDisplacements('ANGULAR Z DISP EXTREME', 14,15);
    end

  else if flexVer1 >= 6 then
    begin
      ReadDisplacements('Table 1.1', 4,5);
      ReadDisplacements('Table 1.2', 6,7);
      ReadDisplacements('Table 1.3', 8,9);
      ReadDisplacements('Table 1.4', 10,11);
      ReadDisplacements('Table 1.5', 12,13);
      ReadDisplacements('Table 1.6', 14,15);
    end

  else
    begin
      Writeln('Flexcom version lower than 8 not supported'); halt(1);
    end;

end;  // Displacements

{**************************************************************************** }
// if APISTD2RD METHOD 4 is require then check validity of ELEMENT PROPERTIES
{ *************************************************************************** }
procedure CheckValidMethod4;

var
  tempstr,tempstr2,tempstr3,dumstr : string;
  i,j,elemno,code : integer;
  endelemprop : boolean;
  firstelemprop : boolean;
  EIyy : real;
begin;
  for i := 1 to TotalElements do
        APISTDValMethod4[i] := false;
  if APISTDMethod4 = true then
  Reset(o);
  dumstr := FindLine(o, '*** ELEMENT PROPERTIES');
  if uppercase(dumstr) <> 'NOT FOUND' then
   begin;
    for j := 1 to 4 do Readln(o);
    readln(o,tempstr);
    endelemprop := false;
    firstelemprop := false;
    repeat;
    readln(o,tempstr);
    Firstword(tempstr,tempstr2);
    Val(tempstr2,elemno,code);
    if (code <> 0) or (elemno = 0) then
      begin;
      if firstelemprop = true then endelemprop := true
      else
      begin;
       readln(o,tempstr);
       Firstword(tempstr,tempstr2);
       Val(tempstr2,elemno,code);
      end;
      end;
    firstelemprop := true;
    Firstword(tempstr,tempstr3);
    Val(tempstr3,EIyy,code);
    if code <> 0 then
     begin;
      for i := 1 to totalelements do
        if elemno = eleminfo[i,1] then APISTDValMethod4[i] := true;
     end;
    until(EOF(o) or (endelemprop = true));
   end;
end;

{**************************************************************************** }
// Indentify if any labels have been used in the Flexcom File
{ **************************************************************************** }
procedure ReadLabels;

var
   tempstr, tempstr2, tempstr3: string;
   endoflabels : boolean;
   i, j, elemi, nodei, code: integer;
   elemno, nodeno: integer;

begin
  Reset(o);
  labeldata := True;
  endoflabels := False;
  dumStr := FindLine(o, '*** LABEL DATA ***');

  if uppercase(dumstr) = 'NOT FOUND' then
    labeldata := False
  else
  begin
    elemi := 0;
    nodei := 0;

    // Skip 4 lines
    for i := 1 to 4 do
      Readln(o);

    repeat
      readln(o,tempstr);
      Firstword(tempstr,tempstr2);

      // Element label
      if Uppercase(tempstr2) = 'ELEMENT' then
      begin
        inc(elemi);

        if elemi > MAX_LABELS then
        begin
          writeln('ERROR - TOO MANY LABELS');
          halt(1);
        end;

        // Extract element number
        Firstword(tempstr,tempstr3);
        Val(Tempstr3,Elemno,Code);

        if code <> 0 then
        begin
          writeln('ERROR in reading label data');
          halt(1);
        end;

        // Store element label and element number
        LabelElem[elemi] := Trim(tempstr);
        LabelElemI[elemi,1] := Elemno;
        LabelElemI[elemi,2] := Elemno;
      end
      // Node label
      else if Uppercase(tempstr2) = 'NODE' then
      begin
        inc(nodei);

        if nodei > MAX_LABELS then
        begin
          writeln('ERROR - TOO MANY LABELS');
          halt(1);
        end;

        // Extract node number
        Firstword(tempstr,tempstr3);
        Val(Tempstr3,Nodeno,Code);

        if code <> 0 then
        begin
          writeln('ERROR in reading label data');
          halt(1);
        end;

        // Store node label and node number
        LabelNode[nodei] := Trim(tempstr);
        LabelNodeI[nodei,1] := nodeno;
        LabelNodeI[nodei,2] := nodeno;
      end
      else
        endoflabels := True;
    until(endoflabels = True);

    TotalLabelE := elemi;
    TotalSingleElemLabels := TotalLabelE;
    TotalLabelN := nodei;

    // Additionally, determine the element range of each line/section by locating
    // each _First _Last element set and storing the first and last element
    for i := 1 to TotalLabelE do
    begin
      // Search for a _First element label
      if (AnsiPos('_First', LabelElem[i]) <> 0) then
      begin
        inc(elemi);

        if elemi > MAX_LABELS then
        begin
          writeln('ERROR - TOO MANY LABELS');
          halt(1);
        end;

        // Store element label without _First and store start element number
        LabelElem[elemi] := AnsiLeftStr(LabelElem[i], AnsiPos('_First', LabelElem[i])-1);
        LabelElemI[elemi,1] := LabelElemI[i,1];
        LabelElemI[elemi,2] := 0;

        // Search for corresponding _Last element label and store last element number
        for j := 1 to TotallabelE do
          if (AnsiPos('_Last', LabelElem[j]) <> 0) then
            if LabelElem[elemi] = AnsiLeftStr(LabelElem[j], AnsiPos('_Last', LabelElem[j])-1) then
              LabelElemI[elemi,2] := LabelElemI[j,2];

        // Error check if last element of line/section not found
        if LabelElemI[elemi,2] = 0 then
        begin
          writeln('Error in reading labels from .out file. Removing ', LabelElem[elemi]);
          elemi := elemi - 1;
        end;
      end;
    end;

    TotalLabelE := elemi;
  end;

  if (TotalLabelN = 0) and (TotalLabelE = 0) then labeldata := False;
end;

{ **************************************************************************** }
// Read geometry data
{ **************************************************************************** }
procedure Geometry;

var
  id, dd, db : Real;
  i, elem, start_nd, end_nd : Integer;

begin
  Reset(o);
  // Find connectivity location
  dumStr := FindLine(o, 'Element Nodal Numbers');

  for elem := 1 to 5 do ReadLn(o);

  // Read and write connectivities
  for i := 1 to totalElements do begin
    ReadLn(o, elem, start_nd, end_nd);
    elemInfo[i,1] := elem;
    elemInfo[i,2] := start_nd;
    elemInfo[i,3] := end_nd;
  end;

  Reset(o);
  // Find section diameters
  dumStr := FindLine(o, '*** DRAG AND BUOYANCY DATA');

  ReadLn(o); ReadLn(o); ReadLn(o); ReadLn(o);
  MaxFluidVal := 0;

  // Determine the number of internal fluid sets
  for i := 1 to totalElements do begin

    If flexVer1 >= 6 then
      begin
        // Read section diameters
        ReadLn(o, elem, id, dd, db, longstr, tempStr);

        if (Pos('None', tempStr) = 0) or (Pos('-',tempstr) = 0) then
          Val(tempStr, fluidVal, code)
        else
          fluidVal := 0;

        id := id + intCoat*2;
        if fluidVal > MaxFluidVal then MaxFluidVal := fluidVal;

      end

    else
      begin
        Writeln(' Flexcom versions below 6 not supported '); halt(1);
      end;

    pipeInfo[i,9] := id;
    elemInfo[i,4] := fluidVal;

  end;

end;  // Geometry


{ **************************************************************************** }
// Calculates effective tensions.
{ **************************************************************************** }
procedure EffectiveTension;

var
  minTen1, maxTen1, minTen3, maxTen3 : Real;
  i, elem, sprCount, flexJntCount : Integer;

begin
  if (flexVer1 >= 8) and (ExtremeData = False) then
    ReadForces('EFFECTIVE TENSION ENVELOPE', 8, ' EFFECTIVE TENS.', '    kN      kN  ')
  else if (flexVer1 >= 8) and (ExtremeData = True) then
    ReadForces('EFFECTIVE TENSION EXTREME', 8, ' EFFECTIVE TENS.', '    kN      kN  ')
  else
    begin;
    Writeln('Flexcom version earlier than 8 not supported'); halt(1);
    end;

  sprCount := 1;
  flexJntCount := 0;

  for i := 1 to totalElements do begin

    elem := elemInfo[i,1];

    if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) then
      begin
        if elem = spring[sprCount] then Inc(sprCount)
        else if flexJntCount < High(flexJoint) then Inc(flexJntCount)
      end

    else
      begin
        minten1 := loads[i,rRef,1,currentFileNo];
        maxten1 := loads[i,rRef+1,1,currentFileNo];
        minten3 := loads[i,rRef,2,currentFileNo];
        maxten3 := loads[i,rRef+1,2,currentFileNo];

        // pipeInfo[i,1] is tension factor, StatTen[i] is static tension  (AG, 27/1/2009)
        minten1 := pipeInfo[i,1]*minten1 + StatTen[i]*(1 - pipeInfo[i, 1]) + tensCorr[i]*1000;
        maxten1 := pipeInfo[i,1]*maxten1 + StatTen[i]*(1 - pipeInfo[i, 1]) + tensCorr[i]*1000;
        minten3 := pipeInfo[i,1]*minten3 + StatTen[i]*(1 - pipeInfo[i, 1]) + tensCorr[i]*1000;
        maxten3 := pipeInfo[i,1]*maxten3 + StatTen[i]*(1 - pipeInfo[i, 1]) + tensCorr[i]*1000;

        loads[i,rRef,1,currentFileNo] := minten1;
        loads[i,rRef+1,1,currentFileNo] := maxten1;
        loads[i,rRef,2,currentFileNo] := minten3;
        loads[i,rRef+1,2,currentFileNo] := maxten3;
      end;

  end;

end; // EffectiveTension


{ **************************************************************************** }
// Calculates pressures.
{ **************************************************************************** }
procedure Pressures;

var
  intpresst, intpresen, extpresst, extpresen, startelev, endelev, dumrel : Real;
  startelevMin, endelevMin,intpresstMin,extpresstMin : real;
  destoppres, desfluidden, desrefelev, intdespresst, intdespresen : Real;
  pminpres, pminfluidden, pminrefelev, intpminst, intpminen : Real;
  intzeropresst, intzeropresen,toppressure : real;
  q : Text;
  i, ErrorCk, elem, sprCount, flexJntCount : Integer;
  // String handling variables
  textDump, dump, temp : String;
  atmpressure,depth,waterdens : real;

begin

  // Finds mud data
  dumStr := FindLine(o, '*** OCEAN ENVIRONMENT DATA');

  for i := 1 to 5 do ReadLn(o);

  atmpressure := 1.01325e5;

  // Values from .out file
  ReadLn(o, depth, waterdens);

  if ExtFluidIndex = False then begin
    for i := 1 to totalelements do
       begin;
       ExtFluid[i,1] := depth;
       ExtFluid[i,2] := waterdens;
       ExtFluid[i,3] := atmpressure;
       end;
  end;

  dumStr := FindLine(o, '*** INTERNAL FLUID DATA');  { finds mud data }

  { number of lines to skip is dependent on version }
      for i := 1 to 4 do ReadLn(o, dumStr);
      temp := Copy(dumStr, 6, 2);
      { Missing Line if No Internal fluid Data given ? }
      if Copy(dumStr, 6, 2) <> 'No' then ReadLn(o);

      if (FlexVer1 >=8) then
         if (FlexVer2 <=3) or (FlexVer2 = 5) then
            ReadLn(o);

  for i := 1 to MaxFluidVal do begin

    // Read in the fluid data, if firstword is blank then exit loop
    ReadLn(o, textDump);
    FirstWord(textDump, dump);
    Val(dump, fluidVal, ErrorCk);

    if errorCk = 0 then begin

      // Read in the data
      FirstWord(textDump, dump);
      Val(dump, FluidHeight, ErrorCk);

      FirstWord(textDump, dump);
      Val(dump, FluidDens, ErrorCk);

      FirstWord(textDump, dump);
      Val(dump, FluidPress, ErrorCk);

      // Apply data to fluid array
      if FluidVal > MAX_FLUIDS then begin; writeln('Number of Fluids in Flexcom file greater than allow in 2HSTRE3D'); halt(1); end;
      fluid[fluidVal,1] := fluidheight;             //values from .out file
      fluid[fluidVal,2] := fluiddens;
      fluid[fluidVal,3] := fluidpress;

      if (IntFluidIndex = True) and (i = FluidSetNo) then
        begin
          fluid[fluidVal,1] := IntFluid[fluidVal,1];    //values from INF file
          fluid[fluidVal,2] := IntFluid[fluidVal,2];
          fluid[fluidVal,3] := IntFluid[fluidVal,3];
        end

    end

  end;

  sprCount := 1;
  flexJntCount := 0;

  // Set default internal fluid values to zero
  for i := 1 to 3 do fluid[0,i] := 0.0;

  for i := 1 to totalElements do begin

    elem := elemInfo[i,1];

    if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) then
      begin
        if elem = spring[sprCount] then Inc(sprCount)
        else if flexJntCount < High(flexJoint) then Inc(flexJntCount)
      end

    else
      begin
        destoppres := DNVpres[i,1]*1E6;
        desfluidden := DNVpres[i,2];
        desrefelev := DNVpres[i,3];
        pminpres := DNVpres[i,4]*1E6;
        pminfluidden := DNVpres[i,5];
        pminrefelev := DNVpres[i,6];
        presscorr := Pinfo[i,1] * 1.0e6;
        fluidVal := eleminfo[i,4];
        nodeheight(i, startelev, endelev, 1);
        nodeheight(i, startelevMin, endelevMin, 3);
        if IntFluidElementOverride = false then
        begin;
        toppressure := (fluid[FluidVal,3] + presscorr);
        intpresst := toppressure + (fluid[FluidVal,2]*GRAVITY*(fluid[FluidVal,1] - startelev));
        intpresen := toppressure + (fluid[FluidVal,2]*GRAVITY*(fluid[FluidVal,1] - endelev));
        intpresstMin := toppressure + (fluid[FluidVal,2]*GRAVITY*(fluid[FluidVal,1] - startelevMin));
        end
        else
        begin;
        toppressure := IntFluidElement[i,3] + presscorr;
        intpresst := toppressure + IntFluidElement[i,2]*GRAVITY*(IntFluidElement[i,1] - startelev);
        intpresen := toppressure + IntFluidElement[i,2]*GRAVITY*(IntFluidElement[i,1] - endelev);
        intpresstMin := toppressure + IntFluidElement[i,2]*GRAVITY*(IntFluidElement[i,1] - startelevMin);
        end;
        intdespresst := destoppres + (desfluidden*GRAVITY*(desrefelev - startelev));
        intdespresen := destoppres + (desfluidden*GRAVITY*(desrefelev - endelev));
        intzeropresst := intpresst - toppressure;
        intzeropresen := intpresen - toppressure;
        intpminst := pminpres + (pminfluidden*GRAVITY*(pminrefelev - startelev));
        intpminen := pminpres + (pminfluidden*GRAVITY*(pminrefelev - endelev));

        { check if pressure is change over the run and warning of limitation}
        if abs(intpresst- intpresstMin)/(abs(intpresst)+10*atmpressure) > 0.05 then
           if (abs(startelevMin- startelev) > 10) and (Warn_Pipe_PressureFlag = false) then
             begin;
             Writeln(WARN_PIPE_PRESSURE + IntToStr(elem));
             Inc(warningsCount);
             warningsArr[warningsCount] := WARN_PIPE_PRESSURE + IntToStr(elem);
             Warn_Pipe_PressureFlag := true;
             end;

        if startelev > ExtFluid[i,1] then
          begin
            extpresst := ExtFluid[i,3];
            extpresen := ExtFluid[i,3];
          end

        else
          begin
            extpresst := ExtFluid[i,3] + ExtFluid[i,2]*GRAVITY*(ExtFluid[i,1] - startelev);
            extpresen := ExtFluid[i,3] + ExtFluid[i,2]*GRAVITY*(ExtFluid[i,1] - endelev);
          end;

        if startelevMin > ExtFluid[i,1] then
          extpresstMin := ExtFluid[i,3]
          else
          extpresstMin := ExtFluid[i,3] + ExtFluid[i,2]*GRAVITY*(ExtFluid[i,1] - StartElevMin);
        if abs(extpresst - extpresstMin)/(abs(extpresst)+10*atmpressure) > 0.05 then
          if (abs(startelevMin - startelev) > 10) and (Warn_Pipe_pressureflag = false) then
            begin;
            Writeln(WARN_PIPE_PRESSURE + IntToStr(elem));
            Inc(warningsCount);
            warningsArr[warningsCount] := WARN_PIPE_PRESSURE + IntToStr(elem);
            Warn_Pipe_PressureFlag := true;
            end;

        Pinfo[i,2] := intpresst/1e6;
        Pinfo[i,3] := extpresst/1e6;
        Pinfo[i,4] := intpresen/1e6;
        Pinfo[i,5] := extpresen/1e6;
        Pinfo[i,6] := intdespresst/1e6;
        Pinfo[i,7] := intdespresen/1e6;
        Pinfo[i,8] := intpminst/1e6;
        Pinfo[i,9] := intpminen/1e6;
        Pinfo[i,10] := toppressure/1e6;
        Pinfo[i,11] := intzeropresst/1E6;
        Pinfo[i,12] := intzeropresen/1E6;

    end;

  end;

end;  // Pressures


{ **************************************************************************** }
// Reads forces.
{ **************************************************************************** }
procedure Forces;

var
  CFN : Integer;

begin
  for CFN := 1 to maxFileNo do begin

    currentFileNo := CFN;
    if currentFileNo = 1 then AssignTabFile(filename);
    if currentFileNo = 2 then AssignTabFile(filenameF);
    if currentFileNo = 3 then AssignTabFile(filenameF2);

    rRef := -1;
    Lheader[1] := '          '; Lheader[2] := '          '; Lheader[3] := 'Ele  Node ';

    // Calculate effective tension
    EffectiveTension;

    Lheader[1] := Lheader[1] + '     PRESSURE     ';
    Lheader[2] := Lheader[2] + ' internal external';
    Lheader[3] := Lheader[3] + '    MPa      MPa  ';
    if currentFileNo = 1 then pressures;

    if (flexVer1 >= 8) and (ExtremeData = False) then
      begin
        ReadForces('SHEAR FORCE Y ENVELOPE', 8, '     SHEAR y    ', '    kN      kN  ');
        ReadForces('SHEAR FORCE Z ENVELOPE', 8, '     SHEAR z    ', '    kN      kN  ');
        ReadForces('TORQUE ENVELOPE', 8, '     TORQUE     ', '   kNm     kNm  ');
        ReadForces('BENDING MOMENT Y ENVELOPE', 8, '  BEND. MOM. y  ', '   kNm     kNm  ');
        ReadForces('BENDING MOMENT Z ENVELOPE', 8, '  BEND. MOM  z  ', '   kNm     kNm  ');
        ReadForces('TOTAL BENDING MOMENT ENVELOPE', 8, ' RES. BEND. MOM.', '   kNm     kNm  ');
        ReadForces('TOTAL CURVATURE ENVELOPE', 8, '   BEND RADIUS  ', '    m       m   ');
      end

    else if (flexVer1 >= 8) and (ExtremeData = True) then
      begin
        ReadForces('SHEAR FORCE Y EXTREME', 8, '     SHEAR y    ', '    kN      kN  ');
        ReadForces('SHEAR FORCE Z EXTREME', 8, '     SHEAR z    ', '    kN      kN  ');
        ReadForces('TORQUE EXTREME', 8, '     TORQUE     ', '   kNm     kNm  ');
        ReadForces('BENDING MOMENT Y EXTREME', 8, '  BEND. MOM. y  ', '   kNm     kNm  ');
        ReadForces('BENDING MOMENT Z EXTREME', 8, '  BEND. MOM  z  ', '   kNm     kNm  ');
        ReadForces('TOTAL BENDING MOMENT EXTREME', 8, ' RES. BEND. MOM.', '   kNm     kNm  ');
        ReadForces('TOTAL CURVATURE EXTREME', 8, '   BEND RADIUS  ', '    m       m   ');
      end

    else
      begin
        writeln(' Flexcom version earlier than 8 not supported');
        halt(1); 
      end;
    
  end; // End number of files loop

  AssignTabFile(filename);

end; // Forces


{ **************************************************************************** }
// Calulates von Mises criteria.
{ **************************************************************************** }
procedure VMCalc(el : integer; vmin, strax, strbm, strhoop, strrad, shrcomb : Real;
                  var vmout : Real);

var
  z : Integer;
  strbmsign, stressax, vm : Real;

begin
  vmout := vmin;           // vmout already equals vmin (CDB 13\10\04)

  // Loop for + and - bm
  for z := 0 to 1 do begin
    strbmsign := (z*2 - 1)*strbm;

    // Calculate principal stress
    stressax := strax + strbmsign;

    vm :=  Sqrt((Sqr(stressax - strhoop) + Sqr(strhoop - strrad)
           + Sqr(strrad - stressax))/2 + 3*Sqr(shrcomb));

    if el = flag then WriteLn(x, 'vm stress', vm);

    // Write the data in the c file if vmout is o (ie first pass) or vm is greater than vmout
    if (vm > vmout) or (vmout = 0) then begin
      vmout:= vm;

        tempel        := el;
        tempstrax     := strax;
        tempstrbmsign := strbmsign;
        tempstrhoop   := strhoop;
        tempstrrad    := strrad;
        tempshrcomb   := shrcomb;
        tempvmout     := vmout;

    end;

  end;

end;  // VMCalc

{ Procedure to pre-write data into main APISTD2RD code check }

Procedure APISTD2RDCheck(intpres,extpres,CheckTen,CheckBM,CheckCur : real; i,j : integer; var APISTD2RDUtil : APISTD2RDArray);
{ Input Corresponding tension and bm, with i element number, and j local element number }
var

diamint,diamout,delta0,yield       : real;
od,id,intcorr,extcorr, wallthick,
Pb,py,pel,pp,pc,ty,my,mp ,secmom   : real;

begin;

 od := pipeInfo[i,5];
 id := pipeInfo[i,9];
 Intcorr := pipeInfo[i,6];
 Extcorr := pipeInfo[i,7];
 intcorr := intcorr/1000; // Convert mm to m
 extcorr := extcorr/1000;
 diamout := od - 2*extcorr;
 diamint := id + 2*intcorr;
 wallthick := (diamout - diamint)/2;
 Delta0 := pipeInfo[i,12];
 YieldStr_Fac := 1.0;
 Yield := pipeinfo[i,4]*1e6;
 if globalFactor then
  begin
  YieldStr_Fac := temperatureFactor[i];
  TensileStr_Fac := temperatureFactorTensile[i];
  end
  else
  begin
  GetYieldStrFac(i);
  end;
 if diamint <= 0 then
   diamint := 0.001;
 Pb := BurstK[i]*(YieldStr_Fac*yield+TensileStr_Fac*pipeInfo[i,10]*1E6)* Ln(diamout/diamint);      //Equ. 1 in API STD 2RD
 Py := 2*YieldStr_Fac*yield*wallThick/diamout;                                                     //Equ. 3 in API STD 2RD
 Pel := 2* pipeInfo[i,11]*power(10.0,6)*power(wallThick/diamout, 3.0)/(1-sqr(pipeInfo[i,13]));     //Equ. 4 in API STR 2RD
 Pp := 2.0*wallThick/diamout*YieldStr_Fac*yield*Alphafab[i];                                       //Equ. 6 in API STD 2RD

 if (Collapse_Type[i] = True) then                                                             //use Equ. 5 in API STR 2RD
    Pc := GetPcFromCubicEqu(Pp,Pel,diamout,wallThick,delta0)
 else                                                                                          //use Equ. 2 in API STD 2RD
    Pc := Py * Pel/sqrt(sqr(Py)+sqr(Pel));
 Ty := YieldStr_Fac*yield * pi*(diamout-wallThick)*wallThick;                    //Equ. 10,11 in API STD 2RD
 secmom := pi/64*(power(diamout,4) - power(diamint,4));
 My := 2*YieldStr_Fac*yield*secmom/(diamout-wallThick);                          //Equ. 12 in API STR 2RD
 Mp := YieldStr_Fac*yield/6*(power(diamout,3)-power(diamint,3));                 //Equ. 13 in API STR 2RD

 APISTD2RDUtil := APISTD2RDCalc(intpres,extpres,Pb,Pc,Mp,Ty,My,CheckTen,CheckBM,CheckCur,YieldStr_Fac*yield,secmom,pipeInfo[i,11],diamout,diamint,i);
 APISTD2RDUtil[6] := CheckTen/Ty;
 APISTD2RDUtil[7] := CheckBm/My;
 if intpres >= extpres then
 APISTD2RDUtil[8] := (intpres-extpres)/pb;
 if extpres > intpres then
 APISTD2RDUtil[8] := (extpres-intpres)/pc;
 APISTD2RDUtil[9] := CheckBM/Mp;

end;



{ **************************************************************************** }
// DNV code check as per DNV-OS-F201 Dynamic Risers, January 2001.
{ **************************************************************************** }
procedure DNVCode;

var
  i, j, k, l : Integer;
  sprCount, flexJntCount, elem : Integer;
  t2 : Real;
  gamSC, gamM, alpU, alpFab, gamF, GamE, GamA, alpC, alpC_min : Real;
  SMYS, SMTS, Emod, Pratio, fytemp, futemp, initoval : Real;
  od, id, intcorr, extcorr, wallratio : Real;
  fy, fu : Real;
  // Pressures
  pb, pc, pel, pp, pmin, pact, pld, pe : Real;
  SolveB, SolveC, SolveD, SolveU, SolveV, SolvePhi, Solve2 : Real;
  Mk, Tk, qh, beta, Mk_min,Tk_min : Real;
  // Forces
  TeF, TeEmin, TeEmax, TeA, MF, MEmin, MEmax, MA : Real;
  Ted, Md : Real;
  Mdarray, Tedarray : Array[1..8] of Real;
  pOver, pCollapse : Real;
  bendFac, tensFac, tensMult, bendMult : Real;
  warnFlag2 : Boolean;
  // Temp DNV check results
  DNVInt_min, DNVInt_max, ptemp : real;

begin
  WriteLn;
  sprCount := 1;
  flexJntCount := 0;

  // Loop elements
  for i := 1 to totalElements do begin

    pWarn2 := false;
    elem := elemInfo[i,1];

    if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) then
      begin
        if elem = spring[sprCount] then Inc(sprCount)
        else if flexJntCount < High(flexJoint) then Inc(flexJntCount)
      end

    else
      begin
        warnFlag2 := True;

        // Decode array data to property variables
        gamSC := DNVFactors[i,1];
        gamM := DNVFactors[i,2];
        alpU := DNVFactors[i,3];
        alpFab := DNVFactors[i,4];
        gamF := DNVfactors[i,5];
        gamE := DNVFactors[i,6];
        gamA := DNVFactors[i,7];
        SMYS := pipeInfo[i,4];
        SMTS := DNVMatSpec[i,1];
        Emod := DNVMatSpec[i,2];
        Pratio := DNVMatSpec[i,3];
        fytemp := DNVMatSpec[i,4];
        futemp := DNVMatSpec[i,5];
        InitOval := DNVMatSpec[i,6];

        intcorr := pipeInfo[i,6]/1000;
        extcorr := pipeInfo[i,7]/1000;
        od := pipeInfo[i,5];
        if DNV_Code_Mod = true then
        od := pipeInfo[i,5] - 2*extcorr;
        id := pipeInfo[i,9];
        t2 := (od - id)/2 - intcorr - extcorr;
        if DNV_Code_Mod = true then
        t2 := (od - id)/2 - intcorr;

        tensFac := pipeInfo[i,1];
        bendFac := pipeInfo[i,2];
        tensMult := DNVMults[i,1];
        bendMult := DNVMults[i,2];

        if t2 < 1e-16 then begin
          WriteLn(' Wall thickness is negative or zero - t2=', t2:7:3, ' OD ', od:7:3, ' ID ', id:7:3);
          Writeln(' For Element :', elem);
          Halt(1);
        end;

        wallratio := od/t2;

        if SMYS > SMTS then begin
          WriteLn(' Tensile strength less than yield - SMYS:', SMYS:8:1, ' SMTS:', SMTS:8:1);
          Halt(1);
        end;

        fy := (SMYS - fytemp)*alpU;
        fu := (SMTS - futemp)*alpU;

        if fy > fu then begin
          WriteLn(' Derated yield greater than tensile strength');
          Halt(1);
        end;

        // Burst resistance
        pb := 2/Sqrt(3)*(2*t2)/(od - t2)*min(fy, fu/1.15);

        // System hoop buckling (collapse)
        pel := 2*Emod/(1 - Sqr(Pratio))*Power((t2/od), 3);
        Pp := 2*(t2/od)*fy*alpfab;

        // Loads at start and end of element considered
        for j := 1 to 2 do begin

          // Pressure
          if j = 1 then begin
            pe := Pinfo[i,3];
            pact := Pinfo[i,2];
            pld := Pinfo[i,6];
            pmin := Pinfo[i,8];
          end;

          if j = 2 then begin
            pe := Pinfo[i,5];
            pact := Pinfo[i,4];
            pld := Pinfo[i,7];
            pmin := Pinfo[i,9];
          end;

          if (pmin > pld) and (pWarn2 = False) then
           begin;
           Writeln(WARN_DNV_PMIN_GT_DESIGN_PRESS + IntToStr(elem));
           Inc(warningsCount);
           warningsArr[warningsCount] := WARN_DNV_PMIN_GT_DESIGN_PRESS + IntToStr(elem);
           ptemp := pld;
           pld := pmin;
           pmin := ptemp;
           pWarn2 := true;
           end;

          if DNVpinput = False then
          begin
              pld := pact; pmin := pact;
          end;

          if (pact - 0.01 > pld) and (pWarn = False) then begin
            WriteLn(WARN_DNV_ACT_PRESS_GT_DESIGN_PRESS);
            Inc(warningsCount);
            warningsArr[warningsCount] := WARN_DNV_ACT_PRESS_GT_DESIGN_PRESS;
            pWarn := True;
          end;

          // Solve for pc - considering splitting to separate subroutine - RIH
          SolveB := -pel;
          SolveC := -(Sqr(pp) + pp*pel*InitOval*od/t2);
          SolveD := pel*Sqr(pp);
          SolveU := 1/3*(-1/3*Sqr(SolveB) + SolveC);
          SolveV := 1/2*(2/27*Power(SolveB, 3) - 1/3*SolveB*SolveC + SolveD);
          Solve2 := -SolveV/Sqrt(-Power(SolveU, 3));

          if (Solve2 >= -1) and (Solve2 < 1) then
            SolvePhi := arccos(Solve2)
          else
            begin
              WriteLn(' Error: has occurred in pressure collapse solver');
              if debug then WriteLn(' Solve2 is ', Solve2);
              Halt(1);
            end;

          pc := -2*Sqrt(-SolveU)*cos(1/3*SolvePhi + 1/3*pi) - 1/3*SolveB;

          if (elem = flag) and (STT = False) then
            WriteLn('pact ', pact:7:3, ' pld ', pld:7:3 , ' pmin ', pmin:7:3, ' pe ',pe:7:3);
          if (elem = flag) and (STT = False) then
            WriteLn('pel ', pel:7:3, ' pp ', pp:7:3 , ' pb ', pb:7:3, ' pc ', pc:7:3);

          // Mk Bending moment resistance and Tk axial force resistance
          // alpC strain hardening and wall thinning parameter
          qh := 0;
          if (pld > pe) then qh := (pld - pe)*2/(Sqrt(3)*pb);

          if wallratio < 15 then beta := 0.4 + qh;

          if (wallratio >= 15) and (wallratio <= 60) then
            beta := (0.4 + qh)*(60 - wallratio)/45;

          if wallratio > 60 then beta := 0;

          alpC := 1 + beta*(fu/fy - 1);

          if alpC > 1.2 then alpC := 1.2;

          // Calculate the zero or pmin strain hardening

          qh := 0;
          if (pmin > pe) then qh := (pmin -pe)*2/(Sqrt(3)*pb);

          if wallratio < 15 then beta := 0.4 + qh;

          if (wallratio >= 15) and (wallratio <= 60) then
            beta := (0.4 + qh)*(60 - wallratio)/45;

          if wallratio > 60 then beta := 0;

          alpC_min := 1 + beta*(fu/fy - 1);

          if alpC_min > 1.2 then alpC_min := 1.2;

          if (elem = flag) and (STT = False) then
            WriteLn(' fy ', fy:7:3 ,' AlpC ', alpC:7:3,' AlpC_min ', alpC_min:7:3, ' od ', od:7:3,
                    ' id ', id:7:3, ' t2 ',t2:7:3);

          Mk := fy*alpC*Sqr(od - t2)*t2*1000;
          Tk := fy*alpC*pi*(od - t2)*t2*1000;
          Mk_min := fy*alpC_min*Sqr(od - t2)*t2*1000;
          Tk_min := fy*alpC_min*pi*(od - t2)*t2*1000;

          if (Mk <= 1e-16) or (Tk <= 1e-16) then begin
            WriteLn(' No resistance for element ', elemInfo[i,1]);
            Halt(1);
          end;

          if maxFileNo = 3 then begin
            TeF := loads[i,2,j,2]/1000;

            // Tension factor taken out owing to change in formulation (AG, 27/1/2009)
            TeEmax := (loads[i,2,j,1]/1000 - TeF)*tensMult;
            TeEmin := (loads[i,1,j,1]/1000 - TeF)*tensMult;

            // Bending factor taken out since loads[] have already been multiplied by it in the procedure 'ReadForces'
            MF := loads[i,14,j,2]/1000;
            MEmax := (loads[i,14,j,1]/1000 - loads[i,14,j,3]/1000)*bendMult + loads[i,14,j,3]/1000 - MF;
            MEmin := (loads[i,13,j,1]/1000 - loads[i,13,j,3]/1000)*bendMult + loads[i,13,j,3]/1000 - MF;

            TeA := 0;
            MA := 0;
          end;

          if maxFileNo = 2 then begin
            TeF := loads[i,2,j,2]/1000;

            // Tension factor taken out owing to change in formulation (AG, 27/1/2009)
            TeEmax := loads[i,2,j,1]/1000 - TeF;
            TeEmin := loads[i,1,j,1]/1000 - TeF;

            MF := loads[i,14,j,2]/1000;
            MEmax := loads[i,14,j,1]/1000 - MF;
            MEmin := loads[i,13,j,1]/1000 - MF;

            if (elem = flag) and (STT = False) then
              WriteLn('TeF :', Tef:7:3, ' TeEmax :', TeEmax:7:3, ' MF ', MF:7:3, ' ME ', MEmax:7:3);

            TeA := 0;
            MA := 0;
          end;

          if maxFileNo = 1 then begin
            TeF := loads[i,2,j,1]/1000;
            MF  := loads[i,14,j,1]/1000;
            TeEMax := 0;
            TeEmin := 0;
            TeA := 0;
            MEmax := 0;
            MEmin := 0;
            MA := 0;
          end;

          // Table 5-2 considerations (Design Load Effects)
          for l := 1 to 8 do begin
            MdArray[l] := 0;
            TedArray[l] := 0;
          end;

          Mdarray[1] := gamF*MF + gamE*MEmax + gamA*MA;
          Tedarray[1] := gamF*TeF + gamE*TeEmax + gamA*TeA;
          Mdarray[2] := gamF*MF + gamE*MEmin + gamA*MA;
          Tedarray[2] := gamF*TeF + gamE*TeEmin + gamA*TeA;

          if (gamF > 0.000000000000001) and (gamE > 0.000000000000001) then begin
            Mdarray[3] := (1/gamF)*MF + gamE*MEmax + gamA*MA;
            Tedarray[3] := (1/gamF)*TeF + gamE*TeEmax + gamA*TeA;
            Mdarray[4] := (1/gamF)*MF + gamE*MEmin + gamA*MA;
            Tedarray[4] := (1/gamF)*TeF + gamE*TeEmin + gamA*TeA;
            Mdarray[5] := gamF*MF + (1/gamE)*MEmax + gamA*MA;
            Tedarray[5] := gamF*TeF + (1/gamE)*TeEmax + gamA*TeA;
            Mdarray[6] := gamF*MF + (1/gamE)*MEmin + gamA*MA;
            Tedarray[6] := gamF*TeF + (1/gamE)*TeEmin + gamA*TeA;
            Mdarray[7] := (1/gamF)*MF + (1/gamE)*MEmax + gamA*MA;
            Tedarray[7] := (1/gamF)*TeF + (1/gamE)*TeEmax + gamA*TeA;
            Mdarray[8] := (1/gamF)*MF + (1/gamE)*MEmin + gamA*MA;
            Tedarray[8] := (1/gamF)*TeF + (1/gamE)*TeEmin + gamA*TeA;
          end;

          Md := 0;
          Ted := 0;

          for l := 1 to 8 do begin
            Md := Max(Mdarray[l], Md);
            if Abs(TedArray[l]) > Abs(Ted) then
               Ted := TedArray[l];
          end;

          if (elem = flag) and (STT = False) then WriteLn('Ted ', Ted:7:3, ' Md ', Md:7:3);
          if (elem = flag) and (STT = False) then WriteLn(' Tk ', Tk:7:3, ' Mk ', Mk:7:3);
          if (elem = flag) and (STT = False) then Writeln('Tk_min ', Tk_min:7:3, ' Mk_min ', Mk_min:7:3);

          // Combined Loading Criteria
          // Pipe member subject to bending moment and effective tension and net internal overpressure
          // Internal Overpressure
          if pld > pe then
            pOver := Sqr((pld - pe)/pb)
          else
            pOver := 0;

          if (Ted < 0) and (WARN_DNV_TE_NEG_FLAG=false) then begin;
                Writeln(WARN_DNV_TE_NEG);
                WARN_DNV_TE_NEG_FLAG := TRUE;
                Inc(warningsCount);
                warningsArr[warningsCount] := WARN_DNV_TE_NEG;
                end;

          if 1 - pOver < 0 then
            begin
              if debug and warnFlag1 and warnFlag2 then begin
                WriteLn(WARN_DNV_PIPE_BURST, elem);
                Inc(warningsCount);
                warningsArr[warningsCount] := WARN_DNV_PIPE_BURST + IntToStr(elem);
                warnFlag2 := False;
                warnTemp := False;
              end;

              DNVInt[i,j] := gamSC*gamM*(Md/Mk*0 + Sqr(Ted/Tk)) + pOver;
            end

          else
            DNVInt[i,j] := gamSC*gamM*(Md/Mk*Sqrt(1 - pOver) + Sqr(Ted/Tk)) + pOver;

          if DNV_Code_Mod = true then
           begin;
           if pmin > pe then
              POver := sqr((pmin-pe)/pb)
           else
              pOver := 0;
           if 1 - pOver < 0 then
            begin
              if debug and warnFlag1 and warnFlag2 then begin
                WriteLn(WARN_DNV_PIPE_BURST, elem);
                Inc(warningsCount);
                warningsArr[warningsCount] := WARN_DNV_PIPE_BURST + IntToStr(elem);
                warnFlag2 := False;
                warnTemp := False;
              end;

              DNVInt_Min := gamSC*gamM*(Md/Mk_min*0 + Sqr(Ted/Tk_min)) + pOver;
            end
           else
             DNVInt_Min := gamSC*gamM*(Md/Mk_min*Sqrt(1 - pOver) + Sqr(Ted/Tk_min)) + pOver;
           DNVInt_max := DNVInt[i,j];
           if debug = true then
             if DNVInt_min > DNVINT_max then writeln('DNV Internal Overpressure come from Pmin strain hardening case for element'+IntToStr(elem));
           if pld >= pe then
           DNVInt[i,j] := Max(DNVInt_min,DNVInt_max)
            else
            DNVInt[i,j] := -1; { set to zero if check if never applicable }
          end;

          if pmin < pe then
            pCollapse := Sqr((pe - pmin)/pc)
          else
            pCollapse := 0;

          // Pipe member subject to bending moment and effective tension and net external overpressure
          if DNV_Code_Mod = true then
             DNVExt[i,j] := Sqr(gamSC*gamM)*(Sqr(Md/Mk_min + Sqr(Ted/Tk_min)) + pCollapse)
          else
             DNVExt[i,j] := sqr(gamSC*gamM)*(Sqr(Md/Mk + Sqr(Ted/Tk)) + pCollapse);
          if DNV_Code_Mod = true then
             if pmin > pe then
              DNVExt[i,j] := -1; { Set to zero if check is never applicable }
          DNVmax[i,j] := Max(DNVInt[i,j], DNVExt[i,j]);

        end;  // End start/end element loop

      end;  // End spring and flexJoint exception

  end;  // End of element loop

  warnFlag1 := warnTemp;

end;  // DNVCode

procedure ReadGRDFile(var PreviousTime, time  :real;var notimetraces : integer);

const
   BendString = 'Total Bending Moment Element';
   TenString  = 'Effective Tension Element';
   ShearString = 'Total Shear Force Element';
   CurvatureString = 'Resultant Curvature Element';
   TorqueString    = 'Torque Element';
var
   i,j,k,k2,valtype : integer;
   dumstr,dump : string;
   dumele, dumnode,elem : integer;
   waveelevation : real;
   sprcount, flexjntcount,ew,ew2 : integer;

begin;

if PreviousTime < 0 then
 begin;
  AssignFile(g, filename + '.grd');
  Reset(g);
  for i := 1 to totalElements do
   for k := 1 to 5 do
    for k2 := 1 to 2 do
      TraceString[i,k,k2] := 0;

  // For version 8 and above, additional information exists at the beginning of
  // the .grd file. Therefore determine location of first time analysis entry
  // based on number of timetraces present.

  if flexVer1 > 7 then begin;
    for j := 1 to 6 do ReadLn(g, dumStr);
    noTimeTraces := StrToInt(dumStr) - 1;
    dumStr := FindLine(g, 'Wave Elevation', False);
    Readln(g,dumstr);
    for j := 1 to NoTimetraces do
     begin;
      readln(g,dumstr);
      valtype := 0;
      if AnsiContainsText(dumstr,tenstring) then valtype := 1;
      if AnsiContainsText(dumstr,bendstring) then valtype := 2;
      if AnsiContainsText(dumstr,curvatureString) then valtype := 3;
      if AnsiContainsText(dumstr,shearstring) then valtype := 4;
      if AnsiContainsText(dumstr,torqueString) then valtype := 5;
      if valtype > 0 then
        begin;
         { Find Element Number }
         repeat;
          FirstWord(dumstr,dump);
          val(dump,dumele,code);
         until(code = 0);
         { Find local node number }
         Repeat;
          FirstWord(dumstr,dump);
          val(dump,dumnode,code);
         Until(code=0);
         for i := 1 to totalElements do
         begin;
          elem := eleminfo[i,1];
          if elem = dumele then
           begin;
           if dumnode = 1 then TraceString[i,valtype,1] := j;
           if dumnode = 3 then TraceString[i,valtype,2] := j;
           end;
          end;
        end;
      readln(g,dumstr);
  end;
  { need to check that for version 8 GRD file that no missing data... }
  sprCount := 1;
  FlexJntCount := 0;
  for i := 1 to TotalElements do
    begin;
      elem := eleminfo[i,1];
      ew := round(elemswitch[i]);
      if (ew = 1) or (ew = 2) then
       ew2 := ew
      else
       begin; ew := 1; ew2:= 2; end;
      if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) or (STTElem[i] = 0) then
      begin
        if elem = spring[sprCount] then Inc(sprCount)
        else if (elem = flexJoint[flexJntCount,0])
              and (flexJntCount < High(flexJoint)) then Inc(flexJntCount)
      end
    else
      begin
      if (TraceString[i,1,ew] = 0) or (TraceString[i,1,ew2] = 0) then
       if STTElem[i] > 0 then
        begin;
        writeln('ERROR Tension timetrace not supplied for element: ',elem);
        if ew = ew2 then writeln('Expected at node location ',ew);
        if ew <> ew2 then writeln('Expecting at node locations ',ew ,',',ew2);
        halt(1);
        end
        else
        STTElem[i] := 0;
       if (TraceString[i,2,ew] = 0) or (TraceString[i,1,ew2] = 0) then
        if STTELem[i] > 0 then
         begin;
         writeln('ERROR Resultant Bending timetrace not supplied for element: ',elem);
         if ew = ew2 then writeln('Expected at node location ',ew);
         if ew <> ew2 then writeln('Expecting at node locations ',ew ,',',ew2);
         halt(1);
         end
         else
         STTElem[i] := 0;
       if (APISTDMethod4 = True) and ((TraceString[i,3,ew] = 0) or (TraceString[i,3,ew2]=0)) then
        if STTElem[i] > 0 then
         begin;
         writeln('ERROR Resultant Curvature Timetrace not supplied for element: ',elem);
         if ew = ew2 then writeln('Expected at node location ',ew);
         if ew <> ew2 then writeln('Expecting at node locations ',ew ,',' ,ew2);
         halt(1);
         end
         else
         STTElem[i] := 0;
       if (BSIShearInput = True) and ((TraceString[i,4,ew] = 0) or (TraceString[i,5,ew] = 0) or (TraceString[i,4,ew2] =0) or (TraceString[i,5,ew2]=0)) then
        if STTElem[i] > 0 then
         begin;
         writeln('ERROR Resultant Shear or Torque not supplied for element: ',elem,' at node location ',ew);
         if ew = ew2 then writeln('Expected at node location ',ew);
         if ew <> ew2 then writeln('Expecting at node locations ',ew ,',',ew2);
         halt(1);
         end
         else
         STTElem[i] := 0;
      end;
    end;

 end;
end;

if not EOF(g) then ReadLn(g, time);
if time < PreviousTime then
 begin;
 Writeln;
 Writeln(' ERROR: Inconsistency in GRD file, expecting ',notimetraces,' timetraces excluding Wave Elevation in STT option');
 Writeln(' NOTE : Spring and Flexjoint Elements should not be included');
 halt(1);
 end;

TraceRead[1] := time;

if not EOF(g) then Read(g, waveelevation);
TraceRead[2] := waveelevation;

for i := 1 to notimetraces do
       if not EOF(g) then Read(g, TraceRead[i+2]);
if not EOF(g) then Readln(g);

if EOF(g) then
 begin;
 Time := -1;
 Close(g);
 end;

end;

{ **************************************************************************** }
// Reads stress timetraces.
{ **************************************************************************** }
procedure STTLoads;

var
  sprCount, flexJntCount, i, j, k, elem, node, noTimeTraces : Integer;
  dummy, id, od, secmom, area, tensFac, bendFac, hoopfac, yield, intcorr, extcorr,
  diamint, diamout, startElev, endElev, intpres1, intpres2,
  intpres, extpres, stressbmavg, time, effTen, resMom, tension, stressax, stress1,
  wallThick, Pb, Py, Pel,Pp, Pc, epsilonB,Ty, My, Mp, M1UtilizCur,
  stress2, temp : Real;
  PreviousTime : real;
  ExTimeTraces,NoElementStressTime,TotalNoTimeStep,OSTTAlength : integer;
  hoopstravg,radialstravg,vmelement1,vmelement2,vmelemmax : real;
  OSTTi,stype,PresFactor,ew,ew1,ew2 : integer;
  ten_end_cap_pressure : real;
  Cdiamout, hoopstrint,hoopstrext,radialstressint,radialstressext,stressbmint,stressbmext,shearres :real;
  radialstressavg,shearcomb,zsec : real;
  minshry,maxshry,minshrz,maxshrz,mintorq,maxtorq, absshry, absshrz,abstorq,rescur : real;
  traceindex : integer;


  ISOUtil    : array[1..4] of Real;
  APISTDUtil : APISTD2RDArray;

begin
  sprCount := 1;
  flexJntCount := 0;

  // Set the Tension Value to tmax, tmin, mmax, mmin to zero.
  for i := 1 to totalElements do begin
    tmax[i] := 0;
    tmin[i] := 0;
    mmax[i] := 0;
    mmin[i] := 0;
  end;

  if DNVCodeCheck then begin
    for i := 1 to totalElements do
      for j := 1 to 2 do begin
        DNVIntMax[i,j] := -1;
        DNVExtMax[i,j] := -1;
      end;
  end;

  for i := 1 to totalElements do
    begin
    // Read properties
    elem := elemInfo[i,1];

    if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) then
      begin
        if elem = spring[sprCount] then Inc(sprCount)
        else if flexJntCount < High(flexJoint) then Inc(flexJntCount)
      end

    else
      begin
        od := pipeInfo[i,5];
        id := pipeInfo[i,9];
        fluidVal := elemInfo[i,4];
        tensFac :=  pipeInfo[i,1];
        bendFac := pipeInfo[i,2];
        hoopfac := pipeInfo[i,3];
        yield := pipeInfo[i,4];
        Intcorr := pipeInfo[i,6];
        Extcorr := pipeInfo[i,7];
        presscorr := Pinfo[i,1];
        presscorr := pressCorr*1.0e6;  // Convert MPa to Pa
        nodeheight(i, startElev, endElev, 1);
        intcorr := intcorr/1000; // Convert mm to m
        extcorr := extcorr/1000;
        diamout := od - 2*extcorr;
        diamint := id + 2*intcorr;
        if diamint >= diamout then begin; writeln('ID > OD for element',Eleminfo[i,1]); halt(1); end;

        STTData[i,1] := diamout;
        STTData[i,2] := diamint;
        STTData[i,3] := od;
        STTData[i,4] := id;
        STTData[i,5] := tensFac;
        STTData[i,6] := bendFac;
        if (elemswitch[i] = 1) then
        begin;
        STTData[i,7] := Pinfo[i,3]*1e6;  // External Pressure at start elev;
        STTData[i,8] := Pinfo[i,2]*1e6; // Internal pressure at start elev;
        STTData[i,9] := Pinfo[i,11]*1e6; // Internal zero pressure at start elev
        end
        else
        begin;
        STTData[i,7] := Pinfo[i,5]*1e6;  // External Pressure at end elev;
        STTData[i,8] := Pinfo[i,4]*1e6; // Internal pressure at end elev;
        STTData[i,9] := Pinfo[i,12]*1e6; // Internal zero pressure at end elev
        end;

        end;

  end;  // End element loop

  // Screen updating added back in by CDB and DL, 20/01/03
  GetXY(ScreenX, ScreenY);
  PreviousTime := -1;

  { Workout how many timetrace are requested }
  ExTimeTraces := 0;
  NoElementStressTime := 0;

  sprCount := 1;
  flexJntCount := 0;
  for i := 1 to totalElements do begin

      elem := elemInfo[i,1];
      if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) or (STTElem[i] = 0) then
        begin
          if elem = spring[sprCount] then Inc(sprCount)
          else if (elem = flexJoint[flexJntCount,0])
              and (flexJntCount < High(flexJoint)) then Inc(flexJntCount)
        end
      else
        begin;
        Inc(ExTimeTraces,2);
        if APISTDMethod4 then inc(ExTimeTraces,1);
        if BSIShearInput then inc(ExTimetraces,2);
        if (OSTTElem[i] = 1) then
         begin;
         Inc(NoElementStressTime,1);
         OSTTData[NoElementStressTime] := elem;
         end;
        end;
  end;

  // Initialise count array
  for i := 1 to totalElements do
   for stype := 1 to 3 do count[i,stype] := 0;
  TotalNoTimeStep := 0;
  OSTTAlength := 100;
  OSTTNoElementStressTime := NoElementStressTime;
  setlength(OSTTArray,NoElementStressTime+1,TotalNoTimeStep+100,19);

  repeat
    sprCount := 1;
    flexJntCount := 0;

    ReadGrdFile(previoustime,time,extimetraces);
    if time < 0 then break;
    Inc(TotalNoTimeStep,1);
    OSTTArray[0,TotalNoTimeStep,1] := time;

    if TotalNoTimeStep + 10 > OSTTAlength then
     begin;
       if NoElementStressTime*TotalNoTimeStep > 500000 then
         begin;
         Writeln(' Size of Output array is massive... suggest less output is requested ');
         end;
       setlength(OSTTArray,NoElementStressTime+1,TotalNoTimeStep+100,19);
       OSTTAlength := TotalNoTimeStep+100;
     end;
     previousTime := time;

    GoToXY(ScreenX, ScreenY);
    Write('                             Stress Timetrace Analysis Time (s): ', time:5:2);

    OSTTi := 0; TraceIndex := 2;
    for i := 1 to totalElements do begin

      elem := elemInfo[i,1];
      OSTTiind[i] := 0;

      if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) or (STTElem[i] = 0) then
        begin
          if elem = spring[sprCount] then Inc(sprCount)
          else if (elem = flexJoint[flexJntCount,0])
              and (flexJntCount < High(flexJoint)) then Inc(flexJntCount)
        end

      else
        begin
          ew := round(elemswitch[i]);
          if (ew = 0) and (Flexver1 < 8) then
           begin;
             Writeln('Force Node must specify the node location');
             halt(1);
           end;
          if ew = 1 then begin; ew1:=1; ew2:=1; end;
          if ew = 0 then begin; ew1:=1; ew2:=2; end;
          if ew = 2 then begin; ew1:=2; ew2:=2; end;
          for ew := ew1 to ew2 do
          begin;
           if (ew = 1) and (ew2<>ew1) then
             begin;
             STTData[i,7] := Pinfo[i,3]*1e6;  // External Pressure at start elev;
             STTData[i,8] := Pinfo[i,2]*1e6; // Internal pressure at start elev;
             STTData[i,9] := Pinfo[i,11]*1e6; // Internal zero pressure at start elev
             end
             else if (ew2<>ew1) then 
             begin;
             STTData[i,7] := Pinfo[i,5]*1e6;  // External Pressure at end elev;
             STTData[i,8] := Pinfo[i,4]*1e6; // Internal pressure at end elev;
             STTData[i,9] := Pinfo[i,12]*1e6; // Internal zero pressure at end elev
             end;
          rescur := -1;
          if FlexVer1 > 7 then
           begin;
           effTen := TraceRead[TraceString[i,1,ew]+2];
           resMom := TraceRead[TraceString[i,2,ew]+2];
           if APISTDMethod4 = true then rescur := TraceRead[TraceString[i,3,ew]+2];
           if BSIShearInput = true then
             begin
             shearres := TraceRead[TraceString[i,4,ew]+2];
             abstorq  := TraceRead[TraceString[i,5,ew]+2];
             end;
           end
           else
           begin
            inc(traceIndex,1);
            effTen := TraceRead[traceindex];
            inc(traceIndex,1);
            resMom := TraceRead[Traceindex];
            if APISTDMethod4 = true then
             begin;
             inc(traceindex,1);
             rescur := TraceRead[Traceindex];
             end;
            if BSIShearInput =true then
             begin;
             inc(traceindex,2);
             shearres := TraceRead[TraceIndex-1];
             abstorq  := TraceRead[TraceIndex];
             end;
           end;
          if OSTTElem[i] = 1 then
          begin;
          inc(OSTTi);
          if ew1 <> ew2 then
           begin;
           writeln(' Output timetrace for element ',elem, ' must be a fixed local node');
           halt(1);
           end;
          OSTTiind[i] := OSTTi;
          OSTTArray[OSTTi,TotalNoTimeStep,2] := effTen;
          OSTTArray[OSTTi,TotalNoTimeStep,1] := ResMom;
          OSTTArray[OSTTi,TotalNoTimeStep,16] := Rescur;
          OSTTArray[OSTTi,TotalNoTimeStep,17] := Shearres;
          OSTTArray[OSTTi,TotalNoTimeStep,18] := abstorq;
          end;
          // stats of tension and bending from run
          if TotalNoTimeStep = 1 then
          begin;
            tmax[i] := effTen;
            tmin[i] := effTen;
            mmax[i] := resMom;
            mmin[i] := resMom;
          end
          else
            begin;
            tmax[i] := max(tmax[i],effTen);
            tmin[i] := min(tmin[i],efften);
            mmax[i] := max(mmax[i],resMom);
            mmin[i] := min(mmin[i],resMom);
            end;
          resMom := STTData[i,6]*resMom;     // Bending factor is multiplied here
          effTen := STTData[i,5]*effTen + StatTen[i]*(1 - STTData[i,5]) + tensCorr[i]*1000; // Tension factor and correction applied here

          loads[i,1,1,1] := effTen; loads[i,2,1,1] := effTen;
          loads[i,1,2,1] := effTen; loads[i,2,2,1] := effTen;
          loads[i,13,1,1] := resMom; loads[i,14,1,1] := resMom;
          loads[i,13,2,1] := resMom; loads[i,14,2,1] := resMom;

          diamout := STTData[i,1];
          diamint := STTData[i,2];
          area := pi/4*(Sqr(STTData[i,1]) - Sqr(STTData[i,2]));
          secmom := pi/64*(Sqr(Sqr(STTData[i,1])) - Sqr(Sqr(STTData[i,2])));
          wallthick := (STTData[i,1] - STTData[i,2])/2;
          if zeroPressureCheck[i] = 0 then
            PresFactor := 8                         // Stress will be calculated for given top pressure only
          else
            PresFactor := 9;                        // Stress will be calculated with and without top pressure

        // Loop on internal pressure
        for k := 8 to PresFactor do begin

             extpres := STTData[i,7];
             intpres := STTData[i,k];
             ten_end_cap_pressure := -(extpres*pi/4*Sqr(diamout)) + (intpres*pi/4*Sqr(diamint));
              // if coating present.
             if coatODflag = true then
               begin;
               Cdiamout := pipeinfo[i,14];
               if Cdiamout < diamout then Cdiamout := diamout;
               ten_end_cap_pressure := -(extpres*pi/4*Sqr(Cdiamout)) + (intpres*pi/4*sqr(diamint));
               end;

             tension := efften + ten_end_cap_pressure;

             hoopstrint := hoopfac*((intpres*(Sqr(diamout) + Sqr(diamint))) - (extpres*2*Sqr(diamout)))
                            /(Sqr(diamout) - Sqr(diamint)); // Thick wall theory
             hoopstrext := hoopfac*((2*intpres*Sqr(diamint)) - (extpres*(Sqr(diamout) + Sqr(diamint))))
                            /(Sqr(diamout) - Sqr(diamint));
             // Note: nominal od is intentionally used in this calculation to get conservative force due to external pressure
             hoopstravg := hoopfac*(intpres - extpres)*diamout*0.5/wallthick - intpres;

             radialstressint := -intpres;
             radialstressext := -extpres;
             radialstressavg := -(extpres*diamout + intpres*diamint)/(diamout + diamint);

              // Axialstress due to tension
             stressax := tension / area;

             // Axial stress due to bending
              // (Note: bending factor already multiplied in procedure 'ReadForces', AG, 27/1/2009)
              stressbmint := 0.5*resMom*diamint/secmom;
              stressbmext := 0.5*resMom*diamout/secmom;
              stressbmavg := 0.25*resMom*(diamout + diamint)/secmom;


            if BSICodeCheck = False then   // API-RP-2RD code check..
             begin;
               vmelemmax := 0;
               shearcomb := 0;
               VMCalc(-1,vmelemmax,stressax,stressbmavg,hoopstravg,
                  radialstressavg,shearcomb,vmelemmax);
               VMCalc(-1,vmelemmax,stressax,-stressbmavg,hoopstravg,
                  radialstressavg,shearcomb,vmelemmax);
             end;

            if BSICodeCheck = true then
             begin;
             // Do BSI code check stuff.
             // Shear stress and torque combination
             // in accordance with BS 8010 : part 3 : 1993 less the 1000 factor which is in error
             minshry := loads[i,3,ew,1];
             maxshry := loads[i,4,ew,1];
             minshrz := loads[i,5,ew,1];
             maxshrz := loads[i,6,ew,1];
             mintorq := loads[i,7,ew,1];
             maxtorq := loads[i,8,ew,1];
             absshry := max(sqr(maxshry),sqr(minshry));
             absshrz := max(sqr(maxshrz),sqr(minshrz));
             if (BSIShearInput = false) then
             abstorq := max(maxtorq,-mintorq);
             if (BSIShearInput = false) then
             shearres := Sqrt(absshry + absshrz);
             zsec := 2*secmom/diamout;
             shearcomb := 0.5*abstorq/zsec + 2*shearres/area;
             vmelemmax := 0.0;
             // On inside of wall
             VMCalc(-1,vmelemmax,stressax,stressbmint,hoopstrint,
                  radialstressint,shearcomb,vmelemmax);
             VMCalc(-1,vmelemmax,stressax,-stressbmint,hoopstrint,
                  radialstressint,shearcomb,vmelemmax);
             // On outside of wall
             VMCalc(-1,vmelemmax,stressax,stressbmext,hoopstrext,
                  radialstressext,shearcomb,vmelemmax);
             VMCalc(-1,vmelemmax,stressax,-stressbmext,hoopstrext,
                  radialstressext,shearcomb,vmelemmax);
              if (WARN_BSI_STTFlag = false) and (BSIShearInput = false) then
               begin;
               Inc(warningsCount);
               warningsArr[warningsCount] := WARN_BSI_STT;
               WARN_BSI_STTFLAG := true;
              end;
             end;

        // This section needs to determine the VM and BSI stress to carry over to ...

            if count[i,1] = 0 then
              begin
                 stressmax[i,1] := vmelemmax/pipeInfo[i,4]/1e6;
                 count[i,1] := 1;
                 Maxstrax[i,ew] := tempstrax;
                 Maxstrbmsign[i,ew] := tempstrbmsign;
                 Maxstrhoop[i,ew] := tempstrhoop;
                 Maxstrrad[i,ew] := tempstrrad;
                 Maxshrcomb[i,ew] := tempshrcomb;
                 Maxvmout[i,ew] := tempvmout;
              end
              else
                 if  vmelemmax/pipeInfo[i,4]/1e6 > stressmax[i,1] then
                   begin;
                   stressmax[i,1] := vmelemmax/pipeInfo[i,4]/1e6;
                   Maxstrax[i,ew] := tempstrax;
                   Maxstrbmsign[i,ew] := tempstrbmsign;
                   Maxstrhoop[i,ew] := tempstrhoop;
                   Maxstrrad[i,ew] := tempstrrad;
                   Maxshrcomb[i,ew] := tempshrcomb;
                   Maxvmout[i,ew] := tempvmout;
                   end;

            if (OSTTElem[i] = 1) then
            begin;
             if k = 8 then OSTTArray[OSTTi,TotalNoTimeStep,3] := vmelemmax/pipeInfo[i,4]/1e6;
             if k = 9 then OSTTArray[OSTTi,TotalNoTimeStep,3] := max(vmelemmax/pipeInfo[i,4]/1e6,OSTTArray[i,TotalNoTimeStep,3]);
            end;

            if APISTD2RDcodeCheck = True then   // Check all methods
                begin
                 { internal pressure, external pressure, tension, bending moment, curvature , i element index, j node index}
                if APISTDmethod4 = False then APISTDVAlMethod4[i] := False;
                APISTD2RDCheck(STTDAta[i,k],STTDATA[i,7],efften,resMom,rescur,i,1,APISTDUtil);
                if count[i,3] = 0 then
                  begin;
                  for stype := 0 to 9 do
                  STD2RDout[i,stype+1] := APISTDUtil[stype];
                  count[i,3] := 1;
                  end
                  else
                   for stype := 0 to 9 do
                      STD2RDOut[i,1+stype] := max(APISTDUtil[stype],STD2RDOut[i,1+stype]);
                if (OSTTElem[i] = 1) then
                begin;
                if k = 8 then
                  begin;
                  for stype := 0 to 5 do
                   OSTTArray[OSTTi,TotalNoTimeStep,10+stype] := APISTDUtil[stype];
                  end;
                if k = 9 then
                  begin;
                  for stype := 0 to 5 do
                    OSTTArray[OSTTi,TotalNoTimeStep,10+stype] := Max(OSTTArray[OSTTi,TotalNoTimeStep,10+stype],APISTDUtil[stype]);
                  end;
                end;


                end;

          if ISOCodeCheck = True then
                begin;
                ISOCWO(effTen, resMom, STTDATA[i,k], STTDATA[i,7], STTData[i,3], wallThick, pipeInfo[i,4]*1e6,
                      pipeInfo[i,10]*1e6, pipeInfo[i,11]*1e6, pipeInfo[i,12],
                      pipeInfo[i,13], ISOUtil);
                if count[i,2] = 0 then
                  begin;
                  for stype := 1 to 4 do stressmax[i,1+stype] := ISOUtil[stype];
                  count[i,2] := 1;
                  end
                  else
                  for stype := 1 to 4 do stressmax[i,1+stype] := max(ISOUtil[stype],Stressmax[i,1+stype]);

                if (OSTTElem[i] = 1) then
                begin;
                 if k = 8 then
                  begin;
                  for stype := 1 to 4 do OSTTArray[OSTTi,TotalNoTimeStep,3+stype] := ISOUtil[stype];
                  end;
                 if k = 9 then
                  begin;
                  for stype := 1 to 4 do
                    OSTTArray[OSTTi,TotalNoTimeStep,3+stype] := Max(OSTTArray[OSTTi,TotalNoTimeStep,3+stype],ISOUtil[Stype]);
                  end;
                end;
              end;

          end;  // End pressures loop
       end; // local elements loop..
    end; // End elements loops
    end;

    if DNVCodeCheck then begin
      DNVCode;
      // Set up the DNVMax of time
      for i := 1 to totalElements do
      begin
        if DNVExt[i,1] > DNVExtMax[i,1] then DNVExtMax[i,1] := DNVExt[i,1];
        if DNVInt[i,1] > DNVIntMax[i,1] then DNVIntMax[i,1] := DNVInt[i,1];
        if DNVExt[i,2] > DNVExtMax[i,2] then DNVExtMax[i,2] := DNVExt[i,2];
        if DNVInt[i,2] > DNVIntMax[i,2] then DNVIntMax[i,2] := DNVInt[i,2];
        if (OSTTElem[i] = 1) and (OSTTiind[i] <> 0) then
           begin;
           if elemswitch[i] = 1 then
            begin;
            OSTTArray[OSTTiind[i],TotalNoTimeStep,8] := DNVExt[i,1];
            OSTTArray[OSTTiind[i],TotalNoTimeStep,9] := DNVInt[i,1];
            end;
           if elemswitch[i] = 2 then
            begin;
            OSTTArray[OSTTiind[i],TotalNoTimeStep,8] := DNVExt[i,2];
            OSTTArray[OSTTiind[i],TotalNoTimeStep,9] := DNVInt[i,2];
            end;
           if elemswitch[i] = 0 then
            begin;
            writeln(' Output timetrace for element ',elem, ' must be a fixed local node');
            halt(1);
            end;
           end;
      end;

    end;


  until (TotalNoTimeStep > 500000); // Timestep loop

  WriteLn;

  { Populate the stress array with maximum obtained from time-domain stress check.
  { APISTD2rd Already done and DNVcal }
  for i := 1 to totalelements do
   begin;
   for j :=1 to 2 do
     StressCheckA[i,j,1] := stressmax[i,1];
   if ISOCodeCheck = True then
     begin;
      for j := 1 to 2 do
       for stype := 1 to 4 do
        StressCheckA[i,j,1+stype] := stressmax[i,1+Stype];
     end;

     loads[i,1,1,1] := tmin[i]; loads[i,2,1,1] := tmax[i];
     loads[i,1,2,1] := tmin[i]; loads[i,2,2,1] := tmax[i];
     loads[i,13,1,1] := mmin[i]; loads[i,14,1,1] := mmax[i];
     loads[i,13,2,1] := mmin[i]; loads[i,14,2,1] := mmax[i];
   end;
  OSTTTotalNoTimeStep := TotalNoTimeStep;

end;  // STTLoads


{ **************************************************************************** }
// Controlling procedure to DNVCode.
{ **************************************************************************** }
procedure DNVCal;

var
  i, j : Integer;

begin

  if debug then begin
    if maxFileNo = 1 then WriteLn(' Stress Check consider functional loading only');
    if maxFileNo = 2 then WriteLn(' Functional And Environmental Loading Considered');
  end;

  if not STT then DNVcode;

  if STT then begin
    for i := 1 to totalElements do
      for j := 1 to 2 do begin
        DNVInt[i,j] := DNVIntMax[i,j];
        DNVExt[i,j] := DNVExtMax[i,j];
        DNVmax[i,j] := max(DNVInt[i,j], DNVExt[i,j]);
      end;
  end;

  // Write output to stress check array
  for i := 1 to totalElements do
      for j := 1 to 2 do
        begin;
        StressCheckA[i,j,6] := DNVMax[i,j];
        StressCheckA[i,j,7] := DNVInt[i,j];
        StressCheckA[i,j,8] := DNVExt[i,j];
        if DNVInt[i,j] < 0 then StressCheckA[i,j,10] := -1
          else
          StressCheckA[i,j,10]  := (DNVint[i,j]);
        if DNVExt[i,j] < 0 then StressCheckA[i,j,11] := -1
          else
          StressCheckA[i,j,11] := (sqrt(DNVExt[i,j]));
        StressCheckA[i,j,9] := max(StressCheckA[i,j,10],StressCheckA[i,j,11]);
        end;

end;  // DNVCal

{ **************************************************************************** }
// Procedure for Calculating of Stress (excluding time domian)
{ **************************************************************************** }
procedure CalStress;

var
  tempVar1, tempVar2, exp1str, exp2str, cof1str, cof2str : String;

  i, j, k, elem, node, sprCount, flexJntCount, len1, len2, lentemp,
  code, PresFactor,stype : Integer;
  Moment_Tension, vecPos: Integer;
  maxten, minten, maxshry, minshry, maxshrz, minshrz, diammean, wallthick,
  endcaparea, maxtorq, mintorq, tempr, bmMax, bmMin, od, secmom, area, id, dd,
  db, tensFac, bendFac, hoopfac, yield, intcorr, extcorr, diamint,
  diamout, startElev, endElev, vmax, pressure, tenmax, tenmin,
  ten_end_cap_pressure, intpres, extpres, radialstressint, radialstressext,
  stressaxmax, stressaxmin, stressbmmaxint, stressbmmaxext, stressbmminint,
  stressbmminext, shearres, shearcomb, vonmax, hoopstrint, hoopstrext, z,
  radialstressavg, hoopstravg, stressbmmaxavg, stressbmminavg,
  Pb, Pc, Py, Pel, Pp, Ty, My, Mp, M_max,epsilonB,
  cof1, cof2, exp1, exp2 : Real;
  Cdiamout : real;
  absten,absbm :real;
  absshry,absshrz,abstorq,abscur : real;

  ISOUtilizationNumber : array[1..4] of Real;
  APISTDUtil : APISTD2RDArray;

begin

  if APISTD2RDcodeCheck then
  begin
      Num_M_Max := 0;
      Num_Bending_Strain := 0;
      for i := 1 to totalElements do
        for stype := 1 to 10 do
          STD2RDOut[i,stype] := 0;
  end;

  sprCount := 1;
  flexJntCount := 0;

  // Read properties
  for i := 1 to totalElements do begin

    elem := elemInfo[i,1];

    if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) then
      begin

        if elem = spring[sprCount] then Inc(sprCount)
        else if flexJntCount < High(flexJoint) then Inc(flexJntCount);

      end

    else
      begin
        od := pipeInfo[i,5]; id := pipeInfo[i,9]; fluidVal := elemInfo[i,4];
        tensFac :=  pipeInfo[i,1]; bendFac := pipeInfo[i,2]; hoopfac := pipeInfo[i,3];
        yield := pipeInfo[i,4]; Intcorr := pipeInfo[i,6]; Extcorr := pipeInfo[i,7];
        presscorr := Pinfo[i,1];
        yield := yield * 1.0e6;     // Converts from MPa to Pa
        presscorr := presscorr * 1.0e6;  // Converts from MPa to Pa

        // Write shutin pressure to file
        pressure := (fluid[FluidVal,3] + presscorr) / 1.0e6;


        intcorr := intcorr/1000;           // Converts from mm to m
        extcorr := extcorr/1000;
        nodeheight(i, startElev, endElev, 1);  // Elem
        diamout := od - 2*extcorr;
        diamint := id + 2*intcorr;
        diammean := (diamout + diamint)/2;
        endcaparea := pi*Sqr(diamint)/4;
        wallThick := (diamout - diamint)/2;

        if wallThick <= 0 then begin
          WriteLn('OD less than or equal to ID');
          WriteLn('Element number is : ', elem);
          Halt(1);
        end;

        area := pi/4*(Sqr(diamout) - Sqr(diamint));
        // Calculate I
        secmom := pi/64*(Sqr(Sqr(diamout)) - Sqr(Sqr(diamint)));

        // Twice, each for start and end node of elem
        // Read forces
        for j := 1 to 2 do begin
          minten := loads[i,1,j,1];
          maxten := loads[i,2,j,1];
          minshry := loads[i,3,j,1];
          maxshry := loads[i,4,j,1];
          minshrz := loads[i,5,j,1];
          maxshrz := loads[i,6,j,1];
          mintorq := loads[i,7,j,1];
          maxtorq := loads[i,8,j,1];
          bmMin   := loads[i,13,j,1];
          bmMax   := loads[i,14,j,1];
          abscur  := loads[i,18,j,1];

          // Reset stress values, vmax
          vmax := 0;

          if zeroPressureCheck[i] = 0 then
            PresFactor := 1                         // Stress will be calculated for given top pressure only
          else
            PresFactor := 0;                        // Stress will be calculated with and without top pressure

          for k := PresFactor to 1 do begin

            // Calculate pressures
            if j = 1 then
              begin
                intpres := Pinfo[i,2]*1e6;
                extpres := Pinfo[i,3]*1e6;
              end

            else
              begin
                intpres := Pinfo[i,4]*1e6;
                extpres := Pinfo[i,5]*1e6;
              end;

              // Tension due to end cap pressure (default to use outer strength diameter
              ten_end_cap_pressure := -(extpres*pi/4*Sqr(diamout)) + (intpres*pi/4*Sqr(diamint));
              // if coating present.
              if coatODflag = true then
               begin;
               Cdiamout := pipeinfo[i,14];
               if Cdiamout < diamout then Cdiamout := diamout;
               ten_end_cap_pressure := -(extpres*pi/4*Sqr(Cdiamout)) + (intpres*pi/4*sqr(diamint));
               end;

              tenmax := maxten + ten_end_cap_pressure;
              tenmin := minten + ten_end_cap_pressure;

              hoopstrint := hoopfac*((intpres*(Sqr(diamout) + Sqr(diamint))) - (extpres*2*Sqr(diamout)))
                            /(Sqr(diamout) - Sqr(diamint)); // Thick wall theory
              hoopstrext := hoopfac*((2*intpres*Sqr(diamint)) - (extpres*(Sqr(od) + Sqr(diamint))))
                            /(Sqr(diamout) - Sqr(diamint));
              // Note: nominal od is intentionally used in this calculation to get conservative force due to external pressure
              hoopstravg := hoopfac*(intpres - extpres)*diamout*0.5/wallthick - intpres;

              radialstressint := -intpres;
              radialstressext := -extpres;
              radialstressavg := -(extpres*diamout + intpres*diamint)/(diamout + diamint);

              // Axialstress due to tension
              stressaxmax := tenmax / area;
              stressaxmin := tenmin / area;

              // Axial stress due to bending
              // (Note: bending factor already multiplied in procedure 'ReadForces', AG, 27/1/2009)
              stressbmmaxint := 0.5*bmMax*diamint/secmom;
              stressbmmaxext := 0.5*bmMax*diamout/secmom;
              stressbmminint := 0.5*bmMin*diamint/secmom;
              stressbmminext := 0.5*bmMin*diamout/secmom;
              stressbmmaxavg := 0.25*bmMax*(diamout + diamint)/secmom;
              stressbmminavg := 0.25*bmMin*(diamout + diamint)/secmom;

              // Shear stress and torque combination
              // in accordance with BS 8010 : part 3 : 1993 less the 1000 factor which is in error

              absshry := max(sqr(maxshry),sqr(minshry));
              absshrz := max(sqr(maxshrz),sqr(minshrz));
              abstorq := max(maxtorq,-mintorq);
              shearres := Sqrt(absshry + absshrz);
              z := 2*secmom/diamout;
              shearcomb := 0.5*abstorq/z + 2*shearres/area;
              // if API-RP-2RD check then shearcomb =0.0 only is BSI check but in Shearcomb.. //

              if elem = flag then begin
                if (k = PresFactor) and (j=1) then
                 begin;
                 AssignFile(x, filename + 'test.nja');   // Element checking file
                 ReWrite(x);
                 end;
                WriteLn(x, 'internal pressure                     :  ', intpres);
                WriteLn(x, 'external pressure                     :  ', extpres);
                WriteLn(x, 'nominal external diameter             :  ', od);
                WriteLn(x, 'nominal internal diameter             :  ', id);
                WriteLn(x, 'max. tension                          :  ', maxten);
                WriteLn(x, 'min. tension                          :  ', minten);
                WriteLn(x, 'tension corrected for pressure        :  ', tenmax);
                WriteLn(x, 'tension corrected for pressure        :  ', tenmin);
                WriteLn(x, 'internal diameter corrected for corr. :  ', diamint);
                WriteLn(x, 'external diameter corrected for corr. :  ', diamout);
                WriteLn(x, 'tension corrected for pressure        :  ', tenmin);
                WriteLn(x, 'hoop factor                           :  ', hoopfac);

                WriteLn(x, 'tension factor                        :  ', tensFac);
                WriteLn(x, 'bending factor                        :  ', bendFac);
                WriteLn(x, 'hoop stress inside wall               :  ', hoopstrint);
                WriteLn(x, 'hoop stress outside wall              :  ', hoopstrext);
                WriteLn(x, 'radial stress inside wall             :  ', radialstressint);
                WriteLn(x, 'radial stress outside wall            :  ', radialstressext);
                WriteLn(x, 'nominal pipe area                     :  ', area);
                WriteLn(x, 'axial stress max                      :  ', stressaxmax);
                WriteLn(x, 'axial stress min                      :  ', stressaxmin);
                WriteLn(x, 'bending moment max                    :  ', bmMax);
                WriteLn(x, 'bending moment min                    :  ', bmMin);
                WriteLn(x, 'second moment of area                 :  ', secmom);
                WriteLn(x, 'bending stress max  int               :  ', stressbmmaxint);
                WriteLn(x, 'bending stress max  ext               :  ', stressbmmaxext);
                WriteLn(x, 'bending stress min  int               :  ', stressbmminint);
                WriteLn(x, 'bending stress min  ext               :  ', stressbmminext);
                WriteLn(x, 'shear in y dir                        :  ', maxshry);
                WriteLn(x, 'shear in z dir                        :  ', maxshrz);
                WriteLn(x, 'resultant shear                       :  ', shearres);
                WriteLn(x, 'section z                             :  ', z);
                WriteLn(x, 'section torque                        :  ', maxtorq);
                WriteLn(x, 'combined shear and torque             :  ', shearcomb);
              end;

              // Reset vonmax (vmax is reset earlier)
              vonmax := 0;

              if ISOCodeCheck = True then
                ISOCWO(maxTen, bmMax, intPres, extPres, od, wallThick, yield,
                      pipeInfo[i,10]*1e6, pipeInfo[i,11]*1e6, pipeInfo[i,12],
                      pipeInfo[i,13], ISOUtilizationNumber);

              if BSIcodecheck = false then  { always do the API-RP-2RD code for a reference stress  }
                begin;
                  // Checks all average stress combinations API-RP-2RD
                  shearcomb := 0.0;
                  VMCalc(elem,vonmax,stressaxmax,stressbmmaxavg,hoopstravg,
                  radialstressavg,shearcomb,vonmax);
                  VMCalc(elem,vonmax,stressaxmax,stressbmminavg,hoopstravg,
                  radialstressavg,shearcomb,vonmax);
                  VMCalc(elem,vonmax,stressaxmin,stressbmmaxavg,hoopstravg,
                  radialstressavg,shearcomb,vonmax);
                  VMCalc(elem,vonmax,stressaxmin,stressbmminavg,hoopstravg,
                  radialstressavg,shearcomb,vonmax);
                end;

              if APISTD2RDcodeCheck = True then
              begin;
                  absten := max(abs(maxten),abs(minten));
                  absbm  := max(abs(bmMax),abs(BmMin));
                  APISTD2RDCheck(intPres,extpres,absten,absbm,abscur,i,j,APISTDUtil);
                  for stype := 0 to 9 do
                    STD2RDOUT[i,stype+1] := Max(STD2RDOUT[i,stype+1],APISTDUtil[stype]);
              end;

              if (BSIcodeCheck = true) then
                begin
                  // Checks all stress combinations:
                  // On inside of wall
                  VMCalc(elem,vonmax,stressaxmax,stressbmmaxint,hoopstrint,
                  radialstressint,shearcomb,vonmax);
                  VMCalc(elem,vonmax,stressaxmax,stressbmminint,hoopstrint,
                  radialstressint,shearcomb,vonmax);
                  VMCalc(elem,vonmax,stressaxmin,stressbmmaxint,hoopstrint,
                  radialstressint,shearcomb,vonmax);
                  VMCalc(elem,vonmax,stressaxmin,stressbmminint,hoopstrint,
                  radialstressint,shearcomb,vonmax);

                  // On outside of wall
                  VMCalc(elem,vonmax,stressaxmax,stressbmmaxext,hoopstrext,
                  radialstressext,shearcomb,vonmax);
                  VMCalc(elem,vonmax,stressaxmax,stressbmminext,hoopstrext,
                  radialstressext,shearcomb,vonmax);
                  VMCalc(elem,vonmax,stressaxmin,stressbmmaxext,hoopstrext,
                  radialstressext,shearcomb,vonmax);
                  VMCalc(elem,vonmax,stressaxmin,stressbmminext,hoopstrext,
                  radialstressext,shearcomb,vonmax);
                end;

              if vonmax > vmax then vmax := vonmax;

            end;  // Pressure loop

            StressCheckA[i,j,1] := vmax/yield;

            if ISOCodeCheck = True then
              begin;
              StressCheckA[i,j,2] := ISOUtilizationNumber[1];
              StressCheckA[i,j,3] := ISOUtilizationNumber[2];
              StressCheckA[i,j,4] := ISOUtilizationNumber[3];
              StressCheckA[i,j,5] := ISOUtilizationNumber[4];
              end;

          // Compare the results from node 1 and 2 of the same element (hence the tempVar1 and 2 and the comparison at the end)
             Maxstrax[i,j]     := tempstrax;
             Maxstrbmsign[i,j] := tempstrbmsign;
             Maxstrhoop[i,j]   := tempstrhoop;
             Maxstrrad[i,j]    := tempstrrad;
             Maxshrcomb[i,j]   := tempshrcomb;
             Maxvmout[i,j]     := tempvmout;

        end;  // Start and end nodes loop

      end;

      if elem = flag then close(x);
  end;  // Element loop



end; // VonMises

function FindSpringValue(Count : integer; SprValue : real) : real;

{ Algorithm below assumes that the force deflection relation is monotonic,
 for the function that is constant by the left boundary of its domain it will assume
 that function is ascending // MP 07.20.02} { Modified to include right boundary! RIH}

var
   i, j : integer;
   force1, force2 : real;
   extension1, extension2 : real;
   ValueFound, MultValue : boolean;
   ExtValue,ExtValue2 : real;

begin;
      ExtValue := 0; { set default to zero in case on value found }
      ValueFound := false; MultValue := false;
      j := 1;
      extension1 := SpringMatData[Count,j,1];
      force1 := SpringMatData[Count,j,2];
      if springdata[Count,3] > 1 then j := 2;
      extension2 := SpringMatData[Count,j,1];
      force2 := SpringMatData[Count,j,2];

      { Does the value exist on the left side of the data ? (Interpolated)}
      if force1 < force2 then
         if SprValue < force1 then
            begin;
            { extrapolate to value }
            ValueFound := true;
            ExtValue := extension1 - (force1 - SprValue)/(force2 - force1)*(extension2-extension1);
            end;
      if force1 > force2 then
         if SprValue > force1 then
            begin;
            { extrapolate to value }
            ValueFound := true;
            ExtValue := extension1 - (SprValue - force1)/(force1 - force2)*(extension2 - extension1);
            end;
      j := 1;
      { Does value exist in central area }
      while (springdata[Count,3] > j) do
          begin;
          extension1 := SpringMatData[Count,j,1];
          force1 := SpringMatData[Count,j,2];
          inc(j,1);
          extension2 := SpringMatData[Count,j,1];
          force2 := SpringMatData[Count,j,2];
          if (force1 >=  force2) then
              if (SprValue > force2) and (SprValue <= force1)  then
                begin;
                if ValueFound = true then
                  begin;
                  if MultValue = false then
                    begin;
                    Inc(warningsCount);
                    warningsArr[warningsCount] := 'Multiple values found for nonlinear spring element ' + inttostr(Spring[Count]) + ' Value: ' + floattostr(SprValue) + ' Extension could be : ' + floattostr(ExtValue);
                    end;
                  MultValue := true;
                  ExtValue2 := ExtValue;
                  end;
                ValueFound := true;
                ExtValue := extension1 + (force1 - SprValue)/(force1 - force2) * (extension2- extension1);
                if MultValue = true then
                    begin;
                    Inc(warningsCount);
                    warningsArr[warningsCount] := 'Multiple values spring ' + inttostr(Spring[Count]) + ' Value: ' + floattostr(SprValue) + ' Extension could be : ' + floattostr(ExtValue);
                    if abs(ExtValue2) < abs(ExtValue) then ExtValue := ExtValue2;
                    end;
                end;
          if (force1 < force2) then
              if (SprValue >= force1) and (SprValue < force2)  then
                begin;
                if ValueFound = true then
                  begin;
                  if MultValue = false then
                    begin;
                    Inc(warningsCount);
                    warningsArr[warningsCount] := 'Multiple values found for nonlinear spring element ' + inttostr(Spring[Count]) + ' Value: ' + floattostr(SprValue) + ' Extension could be : ' + floattostr(ExtValue);
                    end;
                  MultValue := true;
                  ExtValue2 := ExtValue;
                  end;
                ValueFound := true;
                ExtValue := extension1 + (SprValue - force1)/(force2 - force1) *(extension2 - extension1);
                if MultValue = true then
                    begin;
                    Inc(warningsCount);
                    warningsArr[warningsCount] := 'Multiple values spring ' + inttostr(Spring[Count]) + ' Value: ' + floattostr(SprValue) + ' Extension could be : ' + floattostr(ExtValue);
                    if abs(ExtValue2) < abs(ExtValue) then ExtValue := ExtValue2;
                    end;
                end;
          end;
      { Right extrapolation } { force1 and force2 will be the last two values }
      if force1 < force2 then
         if SprValue > force2 then
            begin;
            { extrapolate to value }
            if ValueFound = true then
            begin;
              if MultValue = false then
                begin;
                Inc(warningsCount);
                warningsArr[warningsCount] := 'Multiple values found for nonlinear spring element ' + inttostr(Spring[Count]) + ' Value: ' + floattostr(SprValue) + ' Extension could be : ' + floattostr(ExtValue);
                end;
            MultValue := true;
            ExtValue2 := ExtValue;
            end;
            ValueFound := true;
            ExtValue := extension1 + (force2 - SprValue)/(force2 - force1)*(extension2-extension1);
            if MultValue = true then
              if abs(ExtValue2) < abs(ExtValue) then ExtValue := ExtValue2;
            end;
      if force1 > force2 then
         if SprValue < force2 then
            begin;
            { extrapolate to value }
            if ValueFound = true then
            begin;
              if MultValue = false then
                begin;
                Inc(warningsCount);
                warningsArr[warningsCount] := 'Multiple values found for nonlinear spring element ' + inttostr(Spring[Count]) + ' Value: ' + floattostr(SprValue) + ' Extension could be : ' + floattostr(ExtValue);
                end;
            MultValue := true;
            ExtValue2 := ExtValue;
            end;
            ValueFound := true;
            ExtValue := extension1 + (force2 - SprValue)/(force1 - force2)*(extension2 - extension1);
            if MultValue = true then
              if abs(ExtValue2) < abs(ExtValue) then ExtValue := ExtValue2;
            end;

FindSpringValue := ExtValue;
end;

{ **************************************************************************** }
// Creates spring data table.
{ **************************************************************************** }
procedure Springs;

var
  matlNoStr, dumStr, sprType, tempStr : String;
  i, j, elem, sprCount, temp, code, MaxStrNum : Integer;
  matlNo, sprMax, sprMin, minStroke, maxStroke, stiffness, extension, force,
  extension1, force1,extension2,force2 : Real;
  w, x, y : Text;
  endLoop : Boolean;

begin

  MaxStrNum := 0; 
  for Sprcount := 1 to TotalSprings do
  begin;
    if springdata[Sprcount,1] = 1 then begin
      // Find location
      dumStr := FindLine(o, '*** NONLINEAR MATERIAL TABLES ***');
      MatlNo := springdata[Sprcount,2];
      j := 0;

      if dumStr <> NOT_FOUND then begin

        repeat
          ReadLn(o, dumStr);
          // Find line with matlNo title
          tempStr := Copy(dumStr, 5, 9);
          matlNoStr := Copy(dumStr, 49, 4);
          Val(matlNoStr, temp, code);
          if (Tempstr = 'NONLINEAR') and (Temp > MaxStrNum) then MaxStrNum := temp;

          if EoF(o) then begin
            if MatlNo < 100 then WriteLn('ERROR: Material Properties not found for number ', matlNo);
            if MatlNo < 100 then Halt(1);
            MatlNo := MaxStrNum;
            dumStr := FindLine(o, '*** NONLINEAR MATERIAL TABLES ***');
          end;
        until (tempStr = 'NONLINEAR') and (temp = matlNo);

        for i := 1 to 5 do ReadLn(o);

        if (flexVer1 > 7) or ((flexVer1 = 7) and (flexVer2 >= 9)) then ReadLn(o);

        ReadLn(o, dumStr);
        endLoop := False;

        repeat
          if (flexVer1 > 7) or ((flexVer1 = 7) and (flexVer2 >= 9)) then
            begin
              // Axial Strain
              Val(Copy(dumStr, 3, 11), extension, code);
              // Axial Stress
              Val(Copy(dumStr, 15, 11), force, code);
            end

          else
            begin
              Val(Copy(dumStr, 1, 10), extension, code);
              Val(Copy(dumStr, 12, 10), force, code);
            end;

          inc(j,1);
          SpringData[Sprcount,3] := j;
          if j > Max_Material_Points then
            begin;
            writeln('Number of point of definition of non-linear spring exceed');
            halt(1);
            end;
          SpringMatData[Sprcount,j,1] := extension;
          SpringMatData[Sprcount,j,2] := force;
          ReadLn(o, dumStr);

          if (flexVer1 > 7) or ((flexVer1 = 7) and (flexVer2 >= 9)) then
            if (Copy(dumStr, 4, 3) = '---') then endLoop := True
          else if Length(dumStr) < 20 then endLoop := True;

        until endLoop = True;

      end;

    end;  // End of sprType = 'N'

  end;   // End of list of springs


  SprCount := 1;

  if (flexVer1 >= 8) and (ExtremeData = False) then
    dumStr := FindLine(t, 'EFFECTIVE TENSION ENVELOPE')
  else if (flexVer1 >= 8) and (ExtremeData = True) then
    dumStr := FindLine(t, 'EFFECTIVE TENSION EXTREME')
  else
    begin;
    writeln('Flexcom version below 8 not supported'); halt(1);
    end;

  for i := 1 to 8 do ReadLn(t, dumStr);

  for i := 1 to totalElements do begin

    ReadLn(t, elem, sprMin, sprMax);

    if elem = spring[sprCount] then begin

      if springdata[SprCount,1] = 0 then
        begin
          stiffness := springdata[Sprcount,2];
          if abs(stiffness) < 1E-20 then stiffness := 1E20;
          maxStroke := sprMax/stiffness;
          minStroke := sprMin/stiffness;
        end

      // Nonlinear springs
      // Separate Function to obtain spring values..
      else
      begin;
        maxStroke := FindSpringValue(SprCount,SprMax);
        minStroke := FindSpringValue(SprCount,SprMin);
      end;

      SpringData[SprCount,4] := sprMax;
      SpringData[SprCount,5] := sprMin;
      SpringData[SprCount,6] := MaxStroke;
      SpringData[SprCount,7] := MinStroke;
      inc(sprcount,1);

    end;

  end;

end;  // Springs


{ **************************************************************************** }
// Used to calculate the bending moment or angle of a nonlinear flex joint
// using linear interpolation.
{ **************************************************************************** }
function LinInterp(var xArray, yArray : T2DArray; i : Integer; xp : Real) : Real;

var
  j : Integer;
  x1, x2 : Real;
  inRange : Boolean;

begin
  inRange := False;
  
  for j := 0 to High(xArray[i]) do begin
    x1 := xArray[i,j];
    { if final value }
    if j = High(xArray[i]) then  x2 := x1
    else
    x2 := xArray[i,j+1];

    if (x1 <= xp) and (x2 > xp) then begin
      inRange := True;
      break;
    end;
  end;

  if not inRange then begin
    WriteLn(WARN_LIN_INTERP);
    Inc(warningsCount);
    warningsArr[warningsCount] := WARN_LIN_INTERP;
    j := High(xArray[i])-1;
  end;

  if j < 0 then result := xarray[i,0]
  else 
  Result := yArray[i,j] + (xp - xArray[i,j])*
            (yArray[i,j+1] - yArray[i,j])/(xArray[i,j+1] - xArray[i,j]);

end;  // LinInterp


{ **************************************************************************** }
// Calculates the bending moment or angle of linear and nonlinear flex joints.
{ **************************************************************************** }
procedure FlexJoints;

var
  i, j, k, elem, adjEle1, adjEle2, angEle1, angEle2,
  artNode1, artNode2, node1, node2 : Integer;
  min1, max1, min2, max2, min3, max3,
  bmMax, bmMin, angMax, angMin,
  startNodeMinBM, startNodeMaxBM, endNodeMinBM, endNodeMaxBM : Real;
  w : Text;
  counter : longInt;
  textDump, dump, searchAngleStr, flexJointType, method : String;
  adjElementsFound, angleTimetraceFound, angleInRange,
  zeroStiffness : Boolean;

const
  METHOD_ANGLE = 'Angle Timetrace';
  METHOD_MOMENT = 'BM Envelope';
  FJTYPE_LINEAR = 'Linear';
  FJTYPE_NONLINEAR = 'Nonlinear';

// File definitions:
// t = .tab

begin
  searchAngleStr := '-Angle Timetrace-';
  angleTablesPresent := True;
  angleTimetraceFound := False;
  zeroStiffness := False;

  // Check whether any Angle Timetrace tables exist in .tab file
  textDump := FindLine(t, searchAngleStr);
  if textDump = NOT_FOUND then angleTablesPresent := False;

  // Store header information in array
  SetLength(headerFJArr, 5);
  headerFJArr[0] := 'ARTICULATION / FLEX JOINT DATA';
  headerFJArr[1] := 'Note: Angle Timetrace method is valid for ALL flex joint stiffness';
  headerFJArr[2] := '      Bending Moment method is not valid for LOW flex joint stiffness';
  headerFJArr[3] := 'ELEMENT   FLEX JOINT   METHOD             MAX BM     MIN BM    MAX ANGLE   MIN ANGLE   ANGLE TAKEN BETWEEN';
  headerFJArr[4] := '             TYPE                         (kNm)      (kNm)     (degrees)   (degrees)    ELEMENT & ELEMENT';

  // Set size of flex joint output array (no. of flex joints x no. heading columns)
  SetLength(outputFJArr, totalFlexJoints, 9);
  // Initialise array index of linFJStiffness array
  k := -1;

  //-------------------------
  // Process flex joint data
  //-------------------------
  for i := 0 to totalFlexJoints - 1 do begin

    // Determine whether flex joint is linear or nonlinear
    if flexJoint[i,1] = 0 then
      begin
        flexJointType := FJTYPE_LINEAR;
        Inc(k);
      end
    else if flexJoint[i,1] = 1 then
      flexJointType := FJTYPE_NONLINEAR;

    //------------------------------------
    // Determine flex joint element nodes
    //------------------------------------
    adjElementsFound := True;
    j := 0;
    repeat
      inc(j,1);
      elem := eleminfo[j,1];
    until (elem = flexJoint[i,0]) or (j >= totalElements);
    if elem <> flexJoint[i,0] then
      begin;
      writeln('Error occured flexjoint ', flexJoint[i,0],' not found. ');
      halt(1);
      end;
    artNode1 := eleminfo[j,2];
    artNode2 := eleminfo[j,3];

    // Determine 1st adjacent element
    j := 0;
    adjEle1 := 0;
    adjEle2 := 0;
    repeat
      inc(j,1);
      elem := eleminfo[j,1];
      node1 := eleminfo[j,2];
      node2 := eleminfo[j,3];
      if (node1 = artNode1) or (node2 = artNode1) then
        if (adjele1 = 0) and (elem <> FlexJoint[i,0]) then
              AdjEle1 := elem;
      if (node1 = artNode2) or (node2 = artNode2) then
        if (adjele2 = 0) and (elem <> FlexJoint[i,0]) then
              AdjEle2 := elem;
      // Flag if 1st adjacent element is not found
    until (j >=TotalElements);
    if (AdjEle1 = 0) or (AdjEle2 = 0) then adjElementsFound := false;

    // If flex joint has two adjacent elements
    if adjElementsFound then
      begin

        // Ensure no Angle Timetrace tables are missed in .tab file
        Reset(t);

        //-----------------------------------------------------------
        // Search for Angle Timetrace table pertaining to flex joint
        //-----------------------------------------------------------
        if angleTablesPresent then
          for j := 1 to totalFlexJoints do begin

            angleTimetraceFound := False;

            // Search for angle timetrace table in .tab file
            // (note: .tab file is not reset for each loop)
            textDump := FindLine(t, searchAngleStr, False);

            if textDump <> NOT_FOUND then begin
              ReadLn(t);
              ReadLn(t, textDump);

              // Extract 1st angle element
              for counter := 1 to 5 do FirstWord(textDump, dump);
              Val(dump, angEle1, code);

              // Extract 2nd angle element
              for counter := 1 to 4 do FirstWord(textDump, dump);
              Val(dump, angEle2, code);

              // Check adjacent elements of Angle Timetrace table match
              // adjacent elements of linear flex joint (and exit loop if so)
              if (angEle1 = adjEle1) and (angEle2 = adjEle2) then begin
                angleTimetraceFound := True;
                break;
              end;
            end;

          end;

        //----------------------------------------------------------
        // If Angle Timetrace table found determine bending moments
        //----------------------------------------------------------
        if angleTimetraceFound then
          begin
            method := METHOD_ANGLE;

            // Get max angle
            for counter := 1 to 4 do ReadLn(t, textDump);
            FirstWord(textDump, dump);
            Val(dump, angMax, code);

            // Get min angle
            for counter:= 1 to 2 do FirstWord(textDump, dump);
            Val(dump, angMin, code);

            // Linear flex joint
            if flexJointType = FJTYPE_LINEAR then
              begin
                // Calculate min and max bending moments
                if linFJStiffness[k] <> 0 then
                  begin
                    bmMin := angMin*linFJStiffness[k];
                    bmMax := angMax*linFJStiffness[k];
                  end
                else
                  zeroStiffness := True;
              end

            // Nonlinear flex joint
            else if flexJointType = FJTYPE_NONLINEAR then
              begin
                // Calculate min and max bending moments through linear
                // interpolation of nonlinear flex joint data
                bmMin := LinInterp(nonlinFJTableAng, nonlinFJTableMom, flexJoint[i,2], angMin);
                bmMax := LinInterp(nonlinFJTableAng, nonlinFJTableMom, flexJoint[i,2], angMax);
              end;

          end

        //---------------------------------------------------------
        // Or: Determine angles using Bending Moment Envelope data
        //---------------------------------------------------------
        else
          begin
            method := METHOD_MOMENT;

            // Get bending moment data for adjacent element 1
            if ExtremeData = False then dumStr := FindLine(t, 'TOTAL BENDING MOMENT ENVELOPE');
            if ExtremeData = True  then dumStr := FindLine(t, 'TOTAL BENDING MOMENT EXTREME');
            for j := 1 to 8 do ReadLn(t);

            repeat
              ReadLn(t, elem, min1, max1, min2, max2, min3, max3);
            until elem = adjEle1;

            startNodeMinBM := min3;
            startNodeMaxBM := max3;

            // Get bending moment data for adjacent element 2
            if ExtremeData = False then dumStr := FindLine(t, 'TOTAL BENDING MOMENT ENVELOPE');
            if ExtremeData = True  then dumStr := FindLine(t, 'TOTAL BENDING MOMENT EXTREME');
            for j := 1 to 8 do ReadLn(t);

            repeat
              ReadLn(t, elem, min1, max1, min2, max2, min3, max3);
            until elem = adjEle2;

            endNodeMinBM := min1;
            endNodeMaxBM := max1;

            // Calculate average of bending moments
            bmMin := (startNodeMinBM + endNodeMinBM)/2;
            bmMax := (startNodeMaxBM + endNodeMaxBM)/2;

            // Linear flex joint
            if flexJointType = FJTYPE_LINEAR then
              begin
                // Calculate min and max flex joint angles
                if linFJStiffness[k] <> 0 then
                  begin
                    angMin := bmMin/linFJStiffness[k];
                    angMax := bmMax/linFJStiffness[k];
                  end
                else
                  zeroStiffness := True;
              end

            // Nonlinear flex joint
            else if flexJointType = FJTYPE_NONLINEAR then
              begin
                // Calculate min and max bending moments through linear
                // interpolation of nonlinear flex joint data
                angMin := LinInterp(nonlinFJTableMom, nonlinFJTableAng, flexJoint[i,2], bmMin);
                angMax := LinInterp(nonlinFJTableMom, nonlinFJTableAng, flexJoint[i,2], bmMax);
              end;

          end;

          //--------------------------------------------------------------------
          // Write data to array to be written to file later if bending moments
          // or angles were calculated
          //--------------------------------------------------------------------
          if not zeroStiffness then
            begin
              outputFJArr[i,0] := Format('%5d', [flexJoint[i,0]]) + '     ';
              outputFJArr[i,1] := Format('%-13s', [flexJointType]);
              outputFJArr[i,2] := Format('%-17s', [method]);
              outputFJArr[i,3] := Format('%9.3f', [bmMax/1000]);
              outputFJArr[i,4] := Format('%11.3f', [bmMin/1000]);
              outputFJArr[i,5] := Format('%10.2f', [angMax]);
              outputFJArr[i,6] := Format('%12.2f', [angMin]);
              outputFJArr[i,7] := Format('%11d', [adjEle1]);
              outputFJArr[i,8] := Format('%10d', [adjEle2]);
            end
          // If linear flex joint has zero stiffness, warn that bending moments
          // could not be calculated
          else
            begin
              outputFJArr[i,0] := Format('%5d', [flexJoint[i,0]]) + '     ';
              outputFJArr[i,1] := Format('%-13s', [flexJointType]);
              outputFJArr[i,2] := Format('%-17s', [method]);
              outputFJArr[i,3] := WARN_ZERO_STIFFNESS;
              zeroStiffness := False;
            end;

      end

    // Warn if two adjacent elements do not exist
    else
      begin
        outputFJArr[i,0] := Format('%5d', [flexJoint[i,0]]) + '     ';
        outputFJArr[i,1] := Format('%-13s', [flexJointType]);
        outputFJArr[i,2] := WARN_NO_ADACENT_ELEMENT;
      end;

  end;

end;  // FlexJoints


{ **************************************************************************** }
// Creates the .tb1 file.
{ **************************************************************************** }
procedure CompileTB1;

type
  T1DStrArray = Array[1..3] of String;

var
  r : Text;
  i, j, k, totalLines, len1, len2, code, spaceback, elem,
  sprCount, flexJntCount, vonm1start, vonm2start, vonm1end, vonm2end,
  linectr ,CFN : Integer;
  strTmp : T1DStrArray;
  pageTitle, vm1start, vm1end, vm2start, vm2end, tempStr, tempStr2 : String;
  vonm1, vonm2 : Real;

begin

  // Set up .tb1 file to write to
  if INFnameInOutput = false then
    AssignFile(b, filename + '.tb1')
  else
    AssignFile(b, filename + '_' + infFile + '.tb1');
  ReWrite(b);
  pageTitle := '                 FORCES AND STRESSES      ';
  Header(pageTitle);

  // Write out header lines.
  CFN := 1;

  if ISOCodeCheck then
    begin
      Lheader[1] := Lheader[1] + '  DESIGN DESIGN DESIGN DESIGN';
      Lheader[2] := Lheader[2] + '  FACTOR FACTOR FACTOR FACTOR';
      Lheader[3] := Lheader[3] + '    0.67   0.80   0.90   1.00 ';
    end
  else
    begin
      Lheader[1] := Lheader[1] + '      VON';
      Lheader[2] := Lheader[2] + '   MISES';
      Lheader[3] := Lheader[3] + '    /Yield';
    end;
  if DNVcodeCheck then
    begin;
     Lheader[1] := Lheader[1] +  '   DNV - Code Check   ';
     Lheader[2] := Lheader[2] +  '    Max     Int    Ext  ';
     Lheader[3] := Lheader[3] +  '    Acceptance Value  ';
    end;

  Writeln(b,Lheader[1]);
  WriteLn(b,Lheader[2]);
  WriteLn(b,Lheader[3]);
  sprCount := 1;
  flexJntCount := 0;

  for i := 1 to totalElements do
  begin;
    elem := elemInfo[i,1];
    // determine vonm2 and vonm1 ...
    vonm1 := StressCheckA[i,1,1];
    vonm2 := StressCheckA[i,2,1];
    // do not write spring information
    if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) then
          begin
            if elem = spring[sprCount] then Inc(sprCount)
            else if flexJntCount < High(flexJoint) then Inc(flexJntCount);
          end
        else
          begin
            // determine which internode to output for //
              k := 1;
              if elemSwitch[i] = 1 then k := 1;
              if elemSwitch[i] = 2 then k := 2;
              if (elemSwitch[i] <> 1) and (elemSwitch[i] <> 2) then
                  if vonm2 > vonm1 then k := 2;
              if k =1 then Write(b, elem:3, '    1  ');
              if k =2 then Write(b, elem:3, '    2  ');

              for j := 1 to 16 do begin
                if j < 15 then Write(b, loads[i,j,k,CFN]/1000:7:1, ' ');
                if j >= 15 then Write(b, ' ', loads[i,j,k,CFN]:7:1, ' ');
                if j = 2 then  Write(b, ' ', Pinfo[i,2*k]:6:1, '   ', Pinfo[i,2*k+1]:6:1, ' ');
              end;
              // write stress check values.. 
              if ISOCodeCheck then
                write(b, StressCheckA[i,k,2]:7:3,' ',StressCheckA[i,k,3]:7:3,' ',StressCheckA[i,k,4]:7:3,' ',StressCheckA[i,k,5]:7:3)
              else
                write(b, StressCheckA[i,k,1]:7:3,' ');
              if DNVCodeCheck then
                begin;
                write(b, StressCheckA[i,k,6]:7:3,' ');
                if StressCheckA[i,k,7] < 0 then
                write(b, ' N/A  ')
                else
                write(b,StressCheckA[i,k,7]:7:3,' ');
                if StressCheckA[i,k,8] < 0 then
                write(b, ' N/A  ')
                else
                write(b,StressCheckA[i,k,8]:7:3);
                end;
              WriteLn(b);
          end;
  end; // end element loop

  WriteLn(b);

  pageTitle := '                        FACTORS           ';
  Header(pageTitle);

  write(b, 'Elem Connectivity Tension Bending  Hoop  Yield    Nominal  Nominal  ');
  if (coatodflag = true) then write(b, 'Coating  ');
  write(b,'Internal  External   Tension  ');
  if DNVCodeCheck then
    Write(b,' Gamma   Gamma  Alpha  Alpha  Gamma  Gamma  Gamma  Tensile     Youngs     Poissons   Yield   Tensile  Initial   Enviromental   ');
  if dnvpinput then Write(b, '  Design Presure Inputs                                        ');
  if ISOCodeCheck then Write(b, ' Ultimate   Youngs   Initial   Poisson  ');

  WriteLn(b, ' Shutin ');
  Write(b, '     Node1  Node2 Factor  Factor  Factor Strength   OD        ID        ');
  if (coatodflag = true) then write(b,'OD        ');
  Write(b, 'Corr      Corr   Adjustment');

  if DNVCodeCheck then
    Write(b,'   SC      M      U     FAB     F      E      A    Strength    Modulus      Ratio   Derating Derating Ovality  Bending Tension ');

  if DNVPinput then Write(b, ' Pressure  FluidDen   RefElev    Pressure    FluidDen  RefElev ');

  if ISOCodeCheck then Write(b, ' Strength  Modulus   Ovality    Ratio   ');  //1.51

  WriteLn(b, ' Pressure');
  Write(b, '                                           MPa       m         m         ');
  if (CoatOdflag = true) then write(b,'m         ');
  Write(b, 'mm        mm        kN    ');

  if DNVCodeCheck then
    Write(b, '                                                     MPa         MPa                   MPa       MPa                             ');

  if DNVPinput then Write(b,' MPa       kg/m^3      m        MPa       kg/m^3      m       ');

  if ISOCodeCheck then Write(b, '    MPa      MPa                        ');

  WriteLn(b,'   MPa ');

  for i := 1 to totalElements do
    begin

    Write(b, eleminfo[i,1]:4 ,' ',eleminfo[i,2]:4,' ', eleminfo[i,3]:4,'  '); // element, start_node , end_node
    Write(b, pipeinfo[i,1]:7:3 ,' ',pipeinfo[i,2]:7:3,' ', pipeinfo[i,3]:7:3);// tension, bending and hoop factors
    Write(b, pipeinfo[i,4]:7:1 ,'   ',pipeinfo[i,5]:7:4,'  ', pipeinfo[i,9]:7:4,' '); // Yield Strength , Nominal OD, Nominal ID
    if (coatodflag = true) then write(b, pipeinfo[i,14]:7:3,' '); // coating OD
    write(b,'  ', pipeinfo[i,6]:7:1,'   ', pipeinfo[i,7]:7:1,'    ', pipeinfo[i,8]:7:2,'  '); // Internal Corr , External Corr , Tension Adjustment

    if DNVCodeCheck then
    begin;
    for j := 1 to 7 do Write(b, '  ',DNVfactors[i,j]:4:2, ' ');
    for j := 1 to 6 do Write(b, '  ',DNVMatSpec[i,j]:7:3, ' ');
    for j := 1 to 2 do Write(b, '  ',DNVMults[i,j]:5:3, '  ');

    if DNVpinput then
      for j := 1 to 6 do Write(b, '  ', DNVPres[i,j]:7:2, '  ');
    end;

    if ISOCodeCheck then
    begin;
    Write(b,pipeInfo[i,10]:10:1);
    Write(b,pipeInfo[i,11]:10:1);
    Write(b,pipeInfo[i,12]:10:4);
    Write(b,pipeInfo[i,13]:10:4);
    end;
    // Shut in pressure
    Write(b,'  ' ,Pinfo[i,10]:7:4);
    Writeln(b);
    end;

  WriteLn(b);

  pageTitle := '                      DISPLACEMENTS       ';
  Header(pageTitle);

  Write(b,'              LOCATION         ');
  Writeln(b,' VERTICAL DISP.     HORIZ Y  DISP.     HORIZ Z  DISP.      ANGLE X           ANGLE Y           ANGLE Z  ');
  Write(b,'                               ');
  Writeln(b,' Min      Max       Min      Max       Min      Max      Min      Max      Min      Max      Min      Max  ');
  Write(b,'Node     X         Y         Z ');
  Writeln(b,'  m        m         m        m         m        m      deg      deg      deg      deg      deg      deg');

  for i := 1 to totalNodes do
    begin;

    Write(b, nodeinfoI[i]:3, ' ', nodeinfo[i,1]:8:3, ' ', nodeinfo[i,2]:8:3, ' ', nodeinfo[i,3]:8:3, ' ');
    Write(b, nodeinfo[i,4]:8:3,' ',nodeinfo[i,5]:8:3,' ', nodeinfo[i,6]:8:3,' ',nodeinfo[i,7]:8:3,' ',nodeinfo[i,8]:8:3,' ',nodeinfo[i,9]:8:3);
    Write(b,' ', nodeinfo[i,10]:7:2,'  ',nodeinfo[i,11]:7:2, '  ', nodeinfo[i,12]:7:2, '  ',nodeinfo[i,13]:7:2, '  ',nodeinfo[i,14]:7:2,'  ',nodeinfo[i,15]:7:2);
    WriteLn(b);

    end;

  
  WriteLn(b);

  writeln(b,'                                                                        SPRINGS          ');

  WriteLn(b);
  WriteLn(b, 'SPRING DATA');
  WriteLn(b, 'ELEMENT   MAX FORCE   MIN FORCE   UNITS  MAX STROKE(m)  MIN STROKE(m) (Stroke associated with max and min force)');

  for i := 1 to totalsprings do begin

      Write(b, Spring[i]:6);

      if (Abs(SpringData[i,4]) < 10.0) or (Abs(SpringData[i,5]) < 10.0) then
        WriteLn(b, SpringData[i,4]:12:3, SpringData[i,5]:12:3,'      N', SpringData[i,6]:12:3, SpringData[i,7]:15:3)
      else
        WriteLn(b, SpringData[i,4]:12:1, SpringData[i,5]:12:1,'      N', SpringData[i,6]:12:3, SpringData[i,7]:15:3);
  end;

  WriteLn(b);
  writeln(b,'                                                                        FLEX JOINTS        ');

  // Write header
  for i := 0 to High(headerFJArr) do WriteLn(b, headerFJArr[i]);

  // Write flex joint data
  for i := 0 to High(outputFJArr) do begin
    for j := 0 to 8 do begin

    Write(b, outputFJArr[i,j]);

    end;
    WriteLn(b);
  end;

  // Warn if no Angle Timetrace tables are present in .tab file
  if not angleTablesPresent then begin
    WriteLn(b);
    WriteLn(b, WARN_NO_ANGLE_TIMETRACES);
  end;

  //----------------------------
  // Write a summary of the label data
  //----------------------------
   if labeldata = true then begin;
   writeln(b);
   Writeln(b);

   writeln(b,'                                                                        LABEL INFORMATION');
   writeln(b);
   writeln(b,'                                                                        ELEMENT LABELS');
   writeln(b);
   for i := 1 to totalElements do
    begin;
    write(b,eleminfo[i,1]:4,' ');
    // search for corresponding label for elements
    for j := 1 to TotalLabelE do
      begin;
      if (eleminfo[i,1] >= LabelElemI[j,1]) and (eleminfo[i,1] <= LabelElemI[j,2]) then
        write(b,' {',LabelElem[j],'}  ');
      end;
      writeln(b);
    end;
   writeln(b,'                                                                        NODAL LABELS');
   writeln(b);
   for i := 1 to totalNodes do
    begin;
    write(b,NodeinfoI[i]:4,' ');
    // search for corresponding label for nodes
    for j := 1 to TotalLabelN do
      begin;
      if (NodeinfoI[i] >= LabelNodeI[j,1]) and (NodeInfoI[i] <= LabelNodeI[j,2]) then
        write(b,' {',LabelNode[j],'}  ');
      end;
      writeln(b);
    end;
   writeln(b);
   end; // end of label data

  //----------------------------
  // Write any warnings to file
  //----------------------------
  if warningsCount > 0 then begin
    WriteLn(b);
    Writeln(b,'                                                                   COMPILATION OF WARNINGS  ');

    for i := 1 to warningsCount do begin
      // Write warning from array to .tb1 file
      WriteLn(b, warningsArr[i]);
    end;
  end;


  // Close .tb1 file
  CloseFile(b);

end;  // CompileTB1


{ **************************************************************************** }
// Creates the tb2 file if requested.
{ **************************************************************************** }
procedure CompileTB2;

var
  g : Text;
  i,j, sprCount, flexJntCount, elem : Integer;

begin

  if INFnameInOutput = false then
      AssignFile(g, filename + '.tb2')
  else
      AssignFile(g, filename + '_' + infFile + '.tb2');
  Rewrite(g);

  if STT then
    begin
      WriteLn(g, 'ELEM   AXIAL     BENDING     HOOP      RADIAL     SHEAR      VON MISES    EFFECTIVE TENS  RES BEND MOM');
      WriteLn(g, 'NUM    STRESS    STRESS      STRESS    STRESS     STRESS     STRESS       Min     Max     Min     Max  ');
      WriteLn(g, '       Pa        Pa          Pa        Pa         Pa         Pa            kN      kN     kNm     kNm');
    end

  else
    begin
      WriteLn(g, 'ELEM   AXIAL     BENDING     HOOP      RADIAL     SHEAR      VON MISES');
      WriteLn(g, 'NUM    STRESS    STRESS      STRESS    STRESS     STRESS     STRESS');
      WriteLn(g, '       Pa        Pa          Pa        Pa         Pa         Pa    ');
    end;

  sprCount := 1;
  flexJntCount := 0;

  for i := 1 to totalElements do begin

    // Read element number
    elem := eleminfo[i,1];

    if (elem = spring[sprCount]) or (elem = flexJoint[flexJntCount,0]) then
      begin
        if elem = Spring[sprCount] then Inc(sprCount)
        else if flexJntCount < High(flexJoint) then Inc(flexJntCount)
      end

    else
      begin
        // If is not a spring or flex joint then do the following write data.
        // Compare values and keep data for biggest one
        if Maxvmout[i,2] >= Maxvmout[i,1] then j := 2 else j := 1;

        // .tb2 data
            Write(g, eleminfo[i,1]:6, ' ', formatfloat('0.000E+00',Maxstrax[i,j]):10,
        ' ', formatfloat('0.000E+00', Maxstrbmsign[i,j]):10, ' ', formatfloat('0.000E+00', Maxstrhoop[i,j]):10,
        ' ', formatfloat('0.000E+00', Maxstrrad[i,j]):10, ' ', formatfloat('0.000E+00',Maxshrcomb[i,j]):10,
        ' ', formatfloat('0.000E+00', Maxvmout[i,j]):10);
        if STT then
          // Table format: [Element : Data from tb2 : Statistics]
          Write(g,' ',tmin[i]/1000:7:1, ' ', tmax[i]/1000:7:1, ' ', mmin[i]/1000:7:1, ' ', mmax[i]/1000:7:1, ' ');
      Writeln(g);
      end;

  end;

  CloseFile(g);

end;  // CompileTB2

Procedure CompileTB3;

var
    tt : text;
    NoOutputElem, TotalNoTimeStep : integer;
    i,j,elem : integer;

begin;

    NoOutputElem := OSTTNoElementStressTime;
    TotalNoTimeStep := OSTTTotalNoTimeStep;
    if INFnameInOutput = false then
      assignfile(tt, filename + 'tb3.csv')
    else
      assignfile(tt, filename + '_' + infFile + '_' + 'tb3.csv');
    try
    rewrite(tt);
    except
     begin;
     writeln('Unable to rewrite to ', Filename ,'tb3.csv  might be open in Excel');
     halt(1);
     end;
    end;
    { Write a header to tb3 file }
    WriteLn(tt, 'Program : ',PROG_NAME,'       by 2H Offshore Engineering Limited');
    WriteLn(tt, 'Version : ', VERSION);
    Writeln(tt, 'Filename : ', filename);
    Writeln(tt,'.INF File : ', infFile);
    { Write the explanation header } 
    Writeln(tt, 'LEGEND:');
    if APISTD2RDcodecheck = true then
     begin;
        WriteLn(tt, 'BM        - RESULTANT BENDING MOMENT (kNm)');
        WriteLn(tt, 'TEN       - EFFECTIVE TENSION (kN)');
        Writeln(tt, 'RESCUR    - RESULTANT CURVATURE (1/m)');
        WriteLn(tt, 'M/0.9Mmax     - Bending Moment/90%*Mmax');
        WriteLn(tt, 'BStrain       - Bending Strain');
        WriteLn(tt, 'M1Moment      - utilisation for Method 1');
        WriteLn(tt, 'M2Moment      - utilisation for Method 2');
        WriteLn(tt, 'M3Moment      - utilisation for Method 3');
        WriteLn(tt, 'M4Moment      - utilisation for Method 4');
        WriteLn(tt, 'M4Bstrain     - Bending check for Method 4');
     end
      else
     begin
        WriteLn(tt, 'BM        - RESULTANT BENDING MOMENT (kNm)');
        WriteLn(tt, 'TEN       - EFFECTIVE TENSION (kN)');
        if BSIShearInput = true then
        begin;
        Writeln(tt, 'SHEAR     - Total Shear Force (kN)');
        writeln(tt, 'TOR       - Torque (kNm)');
        end;
        WriteLn(tt, 'VM/YIELD  - VON MISES STRESS TO YIELD STRESS RATIO');
     end;
        WriteLn(tt);
        Write(tt, 'TIME(s)');

     for i := 1 to NoOutputElem do
        begin;
        elem := OSTTData[i];
        Write(tt, ', ', 'BM-', elem, ', ', 'TEN-', elem);
        if APISTDMethod4 = true then
        Write(tt, ',', 'RESCUR-', elem);
        if BSIShearInput = true then
        Write(tt, ', ', 'SHEAR-', elem, ', ', 'TOR-', elem);
        if ISOCodeCheck = true then
            Write(tt, ', ', 'ISO -0.67 ', elem , ',','ISO - 0.80 ', elem ,',' ,'ISO - 0.90', elem ,',','ISO - 1.00', elem );
        if (DNVCodeCheck = true) and (APISTD2RDcodeCheck = false) then
            Write(tt, ', ', 'DNV Ext', elem ,',', ' DNV Int', elem);
        if APISTD2RDcodeCheck = True then
            begin;
            Write(tt, ', ', 'M/0.9Mmax', elem, ', ', '     BStrain', elem, ', ', 'M1Method', elem, ', ', 'M2Method', elem);
            Write(tt, ', ', 'M3Method-Ext', elem, ', ', 'M3Method-Int', elem);
            if APISTDMethod4 = true then write(tt,', ',  'M4Moment', elem, ', ', 'M4Bstrain', elem);
            end;
            Write(tt, ', ', 'VM/YIELD-', elem);
        end;

     writeln(tt);

     for j := 1 to TotalNoTimeStep do
      begin;
      write(tt, OSTTArray[0,j,1]:6:1);
      for i := 1 to NoOutputElem do
       begin;
         Write(tt, ', ', OSTTArray[i,j,1]/1000:10:3, ', ', OSTTArray[i,j,2]/1000:10:3);
         if APISTDMethod4 = true then
         Write(tt, ',', OSTTArray[i,j,16]:8:6);
         if BSIShearInput = true then
         Write(tt, ', ', OSTTArray[i,j,17]/1000:10:3 , ',' , OSTTArray[i,j,18]/1000:10:3);
         if ISOCodeCheck = true then
            begin;
            Write(tt, ',    ', OSTTArray[i,j,4]:8:6, ',      ', OSTTArray[i,j,5]:8:6);
            Write(tt, ',    ', OSTTArray[i,j,6]:8:6, ',      ', OSTTArray[i,j,7]:8:6);
            end;
         if (DNVCodeCheck = true) and (APISTD2RDcodeCheck = false) then
            begin;
            Write(tt, ',    ', OSTTArray[i,j,8]:8:6, ',    ', OSTTArray[i,j,9]:8:6);
            end;
         if APISTD2RDcodeCheck = True then
            begin;
            Write(tt, ',    ', OSTTArray[i,j,10]:8:6, ',      ', OSTTArray[i,j,15]:8:6, ',        ', OSTTArray[i,j,11]:8:6 );
            Write(tt, ',    ', OSTTArray[i,j,12]:8:6, ',      ', OSTTArray[i,j,8]:8:6, ',     ', OSTTARray[i,j,9]:8:6);
            if APISTDMethod4 = true then write(tt, ',    ', OSTTArray[i,j,13]:8:6, ',      ', OSTTArray[i,j,14]:8:6);
            end;
            Write(tt,  ', ', OSTTArray[i,j,3]:10:3);
       end;
      writeln(tt);
      end;

      Close(tt);

end;



{ **************************************************************************** }
// Creates the tb4 & tb5 files if requested.
{ **************************************************************************** }
Procedure CompileTB4TB5;

var
      tb4Out, tb5Out: Text;
      DNVutlInt,DNVutlExt,DNVutlExtSR :real;
      i : integer;
      BendRdMin : real;
begin
      if INFnameInOutput = false then
       begin;
       AssignFile(tb4Out, filename + '.tb4');
       AssignFile(tb5Out, filename + '.tb5');
       end
       else
       begin;
       AssignFile(tb4Out, filename + '_' + infFile + '.tb4');
       AssignFile(tb5Out, filename + '_' + infFile + '.tb5');
       end;
      Rewrite(tb4Out);
      Rewrite(tb5Out);
      // Write out header information
      WriteLn(tb4Out,'Program : ', PROG_NAME);
      WriteLn(tb4Out, 'Version : ', VERSION);
      WriteLn(tb4Out, 'Title   : ', title  );
      WriteLn(tb4Out, '(1) Utilization per DNV-OS-F201 for Internal Overpressure using LRFD method');
      WriteLn(tb4Out, '(2) Utilization per DNV-OS-F201 for External Overpressure using LRFD method');
      WriteLn(tb4Out, '(3) Utilization per API-STD-2RD Method 4, equation 24');
      WriteLn(tb4Out, '(4) Utilization per API-STD-2RD Method 4, equation 26 or 27');
      Write(tb4Out, 'Warning : ');

      // Write out header information
      WriteLn(tb5Out,'Program : ', PROG_NAME);
      WriteLn(tb5Out, 'Version : ', VERSION);
      WriteLn(tb5Out, 'Title   : ', Title);
      if STT = true then
      begin;
      WriteLn(tb5Out, 'Note: When STT option is active for the element, TensionMax and MomentMax are not the maximum value of Tension and Bending Moment.');
      WriteLn(tb5Out, '      TensionMax and MomentMax are the values when maximum utilization occurs based on Method 1 of API-STD-2RD.');
      end;

      if Num_M_Max > 0 then
          WriteLn(tb4Out, Num_M_Max,' (', NOTE_API_STD_2RD_M_MAX, ')')
      else
          WriteLn(tb4Out, 'None');

      if Num_Bending_Strain > 0 then
          WriteLn(tb4Out, '          ', Num_Bending_Strain,' (', NOTE_API_STD_2RD_BSTRAIN, ')')
      else
          WriteLn(tb4Out, '          None');

      WriteLn(tb4Out, 'Element, BM/0.9Mmax, M1Utilization, M2Utilization, M3UtilInt(1), M3UtilExt(2), M4LoadUtil(3), M4StrainUtil(4)');
      WriteLn(tb5Out,'Element,  TensionMax/Ty,  MomentMax/My,  PressureRatio,  MomentMax/Mp,   BendRdMin,    BendStrainMax');
      for i := 1 to totalElements do
      begin
          DNVutlInt := max(DNVInt[i,1], DNVInt[i,2]);
          DNVutlExt := max(DNVExt[i,1], DNVExt[i,2]);
          BendRdMin := min(loads[i,15,1,1],loads[i,15,2,1]);
          if DNVutlExt >= 0 then DNVUtlExtSR := sqrt(DNVutlExt)
          else
          DNVUtlExtSR := -1;
          WriteLn(tb5Out,elemInfo[i,1]:4, ',     ', STD2RDOut[i,7]:8:6, ',        ', STD2RDOut[i,8]:8:6, ',       ', STD2RDOut[i,9]:8:6, ',       ', STD2RDOut[i,10]:8:6, ',    ', BendRdMin:12, ',    ', STD2RDOut[i,6]:8:6);
          if STD2RDOut[i,1] < 1E8 then
          Write(tb4Out,elemInfo[i,1]:4, ',      ', STD2RDOut[i,1]:8:6, ',     ', STD2RDOut[i,2]:8:6, ',   ', STD2RDOut[i,3]:8:6)
          else
          Write(tb4Out,elemInfo[i,1]:4, ',      ', 'Fail MMax=0', ',     ', STD2RDOut[i,2]:8:6, ',   ', STD2RDOut[i,3]:8:6);
          if DNVutlInt < 0 then
          Write(tb4Out,',      ', 'N/A')
          else
          Write(tb4Out,',      ', DNVutlInt:8:6);
          if DNVutlExtSR < 0 then
          Write(tb4out,',      ', 'N/A')
          else
          Write(tb4Out,',     ', DNVutlExtSR:8:6);
          if APISTDValMethod4[i] = true then
          WriteLn(tb4Out,',     ', STD2RDOut[i,4]:8:6, ',     ', STD2RDOut[i,5]:8:6)
          else
          WriteLn(tb4Out,',     ', 'Insufficient Data ' , ',     ', 'Insufficient Data');
      end;

      CloseFile(tb4Out);
      CloseFile(tb5Out);

end;

// add header into data array //
Procedure AddHeader(strrow1, Strrow2 , strrow3 : string; rowstart, headcol : integer);
begin
  if strrow1 <> 'BLANK' then HeaderToExcel[rowstart,headcol] := Strrow1;
  if strrow2 <> 'BLANK' then HeaderToExcel[rowstart+1,headcol] := Strrow2;
  if strrow3 <> 'BLANK' then HeaderToExcel[rowstart+2,headcol] := Strrow3;
end;


Procedure OutputExcel1; { procedure to output data to excel files... }

var
   Dirlist : string;
   TemplateFileName,TemplateLocation,CurrentDirectory : string;
   OutputFileName : string;
   Sheet, Excel : OLEVariant;
   TransferArr,TransferArr2 : oleVariant;
   i,ish,j,j2,k, numWorkSheets,countrow : integer;
   sprCount,flexjntcount,elem : integer;
   k2,CFN,countlab,previouslab,endlab,firstlab : integer;
   different : boolean;
   vonm1,vonm2,tempmin,tempmax,tempmin1,tempmin2,tempmax1,tempmax2 : real;
   DNVutlInt,DNVutlExt,DNVutlExtSR :real;
   TempStr,tempstring1,tempstring2 : string;

begin
   DirList := CurrentDirProgram + ';' + CurrentDirProgram +'Templates\';
{ Search either location of file for template or subfolder template }
   writeln;
   TemplateFileName := 'Template2HSTRE3D.xlsm';
   TemplateLocation := FileSearch(TemplateFileName,DirList);
   if TemplateLocation = '' then
      begin;
      writeln('Template File Missing no excel written');
      writeln('Directory Search List for Template file is ', Dirlist);
      exit;
   end;
   TemplateLocation := ExpandFileName(TemplateLocation);
   CurrentDirectory := GetCurrentDir;
   if INFnameInOutput = false then
      OutputFileName := CurrentDirectory + '\' + filename + '_2HSTRE3D.xlsm'
   else
      OutputFileName := CurrentDirectory + '\' + filename + '_' + infFile + '_2HSTRE3D.xlsm';

{  required from a console application program when using OLE }
  try
     Excel := CreateOleObject('Excel.Application');

  if VarIsEmpty(Excel) = false then
     begin;
     Excel.Visible := false; ///false;
     Excel.IgnoreRemoteRequests := True;
     if FileExists(TemplateLocation) = true then
      Excel.Workbooks.Open(TemplateLocation, ReadOnly := True)
      else
        begin;
        writeln('Template file not found'); exit;
        end;
     { determine number of worksheets }
     numWorkSheets := Excel.Worksheets.Count;
    for ish := 1 to numWorkSheets do
      begin;
      Sheet := Excel.Worksheets[ish];
      Tempstr := Sheet.Name;
      if LowerCase(Tempstr) = 'cover' then
       begin;
       Sheet.Cells[4,3].Value := 'BSI';
       if ISOCodeCheck = True then Sheet.Cells[4,3].Value := 'ISO';  // Stress Check Type
       if APIRP2RDCodeCheck = True then Sheet.Cells[4,3].Value := 'API-RP-2RD';
       if DNVCodeCheck = True then Sheet.Cells[4,3].Value := 'DNV-OS-F201';
       if APISTD2RDCodeCheck = True then Sheet.Cells[4,3].Value := 'API-STD-2RD';
       Sheet.Cells[5,3].Value := WarningsCount;
       i := 6;
       if Functional = true then begin; Sheet.Cells[i,2].Value := 'DNV-OS-F201 code with functional loads only'; inc(i); end;
       if STT = true then begin; Sheet.Cells[i,2].Value := 'STT is turned ON'; inc(i); end;
       if OSTT = true then begin; Sheet.Cells[i,2].Value := 'OSTT is turned ON';  inc(i); end;
       if DNV_Code_Mod = false then begin; Sheet.Cells[i,2].Value := 'Following DNV-OS-F201 Code to letter'; inc(i); end;
       if CoatODFlag = true then begin; Sheet.Cells[i,2].Value := 'Coating OD specified for VM check'; inc(i); end;
       if Debug = true then begin; Sheet.Cells[i,2].Value := 'Debug is turned ON'; inc(i); end;
       end;
      if LowerCase(Tempstr) = 'echodata' then
       begin;
       setlength(HeaderToExcel,10,51);
       HeaderToExcel[0,1] := 'Created by : ' + PROG_NAME + '       by 2H Offshore Engineering';
       HeaderToExcel[1,1] := 'Version    : ' + Version;
       HeaderToExcel[2,1] := 'Title      : ' + Title;
       HeaderToExcel[3,1] := 'Filename   : ' + filename;
       HeaderToExcel[3,5] := '.INF File  : ' + infFile;
       if MaxFileNo > 1 then HeaderToExcel[3,8] := 'Functional Filename : ' + filenameF;
       if MaxFileNo > 2 then HeaderToExcel[3,11]:= 'Environmental Filename 2 : ' + FilenameF2;
       HeaderToExcel[4,5] := ' INPUT FACTORS';
       k := 0;
       AddHeader('Element','BLANK','BLANK',5,k); inc(k);
       AddHeader('Connectivity','Node 1','BLANK',5,k); inc(k);
       AddHeader('BLANK','Node 2','BLANK',5,k); inc(k);
       AddHeader('Tension','Factor','BLANK',5,k); inc(k);
       AddHeader('Bending','Factor','BLANK',5,k); inc(k);
       AddHeader('Hoop','Factor','BLANK',5,k); inc(k);
       AddHeader('Yield','Strength','MPa',5,k); inc(k);
       AddHeader('Nominal','OD','m',5,k); inc(k);
       AddHeader('Nominal','ID','m',5,k); inc(k);
       if coatodflag = true then begin; AddHeader('Coating','OD','m',5,k); inc(k); end;
       AddHeader('Internal','Corrosion','mm',5,k); inc(k);
       AddHeader('External','Corrosion','mm',5,k); inc(k);
       AddHeader('Tension','Adjustment','kN',5,k); inc(k);
       AddHeader('Shutin','Pressure','MPa',5,k); inc(k);
       AddHeader('External Fluid','Elevation','m',5,k); inc(k);
       AddHeader('BLANK','Density','kg/m^3',5,k); inc(k);
       AddHeader('BLANK','Pressure at Reference','Pa',5,k); inc(k);
       AddHeader('Internal Fluid','Elevation','m',5,k); inc(k);
       AddHeader('BLANK','Density','kg/m^3',5,k); inc(k);
       AddHeader('BLANK','Pressure at Reference','Pa',5,k); inc(k);

       if DNVcodeCheck then
        begin;
        AddHeader('Gamma','SC'   ,'BLANK',5,k); inc(k);
        AddHeader('Gamma','M','BLANK',5,k); inc(k);
        AddHeader('Alpha','U','BLANK',5,k); inc(k);
        AddHeader('Alpha','FAB','BLANK',5,k); inc(k);
        AddHeader('Gamma','F','BLANK',5,k); inc(k);
        AddHeader('Gamma','E','BLANK',5,k); inc(k);
        AddHeader('Gamma','A','BLANK',5,k); inc(k);
        AddHeader('Tensile','Strength','MPa',5,k); inc(k);
        AddHeader('Youngs','Modulus','MPa',5,k); inc(k);
        AddHeader('Poissons','Ratio','BLANK',5,k); inc(k);
        AddHeader('Yield','Derating','MPa',5,k); inc(k);
        AddHeader('Tensile','Derating','MPa',5,k); inc(k);
        AddHeader('Initial','Ovality','BLANK',5,k); inc(k);
        AddHeader('Enviromental','Bending','Factor',5,k); inc(k);
        AddHeader('BLANK','Tension','Factor',5,k); inc(k);
        end;
       if DNVPInput then
        begin;
        AddHeader('Maximum Pressure Inputs','Pressure','MPa',5,k); inc(k);
        AddHeader('BLANK','Fluid Density','kg/m^3',5,k); inc(k);
        AddHeader('BLANK','Ref Elev','m',5,k); inc(k);
        AddHeader('Minimum Pressure Inputs','Pressure','MPa',5,k); inc(k);
        AddHeader('BLANK','Fluid Density','kg/m^3',5,k); inc(k);
        AddHeader('BLANK','Ref Elev','m',5,k); inc(k);
        end;
       if ISOCodeCheck then
        begin;
        AddHeader('Ultimate','Strength','MPa',5,k); inc(k);
        AddHeader('Youngs','Modulus','MPa',5,k); inc(k);
        AddHeader('Initial','Ovality','BLANK',5,k); inc(k);
        AddHeader('Poisson','Ratio','BLANK',5,k); inc(k);
        end;
       if APISTD2RDCodeCheck = true then
        begin;
        AddHeader('Collapse','Equation','BLANK',5,k); inc(k);
        AddHeader('Strain','Amplification','Factor',5,k); inc(k);
        AddHeader('Mechanical','Variability','k',5,k); inc(k);
        AddHeader('Collapse','Factor','fc',5,k); inc(k);
        if  globalfactor = false then
          begin;
          AddHeader('Temperature','BLANK','Deg C',5,k); inc(k);
          AddHeader('Thermal','Table','No',5,k); inc(k);
          end
          else
          begin;
          AddHeader('Temperature','Factor','Yield',5,k); inc(k);
          AddHeader('Temperature','Factor','Tensile',5,k); inc(k);
          end;
        end;
       { Output to Excel }
       TransferArr2 := HeaderToExcel;
       Sheet.Range[Sheet.Cells[1,1],Sheet.Cells[9,1+k]] := TransferArr2;
       Finalize(HeaderToExcel); Finalize(TransferArr2);

       setlength(DataToExcel,TotalElements+1,51);
       for i := 1 to TotalElements do
        begin;
        k := 0;
        DataToExcel[i-1,k] := eleminfo[i,1]; inc(k); //element
        DataToExcel[i-1,k] := eleminfo[i,2]; inc(k); //start_node
        DataToExcel[i-1,k] := eleminfo[i,3]; inc(k); //End_node
        DataToExcel[i-1,k] := pipeinfo[i,1]; inc(k); //tension factor
        DataToExcel[i-1,k] := pipeinfo[i,2]; inc(k); //bending factor
        DataToExcel[i-1,k] := pipeinfo[i,3]; inc(k); //hoop factor
        DataToExcel[i-1,k] := pipeinfo[i,4]; inc(k); //yield strength
        DataToExcel[i-1,k] := pipeinfo[i,5]; inc(k); //nominal OD
        DataToExcel[i-1,k] := pipeinfo[i,9]; inc(k); //nominal ID
        if coatodflag = true then begin; DataToExcel[i,k] := pipeinfo[i,14]; inc(k); end;
        DataToExcel[i-1,k] := pipeinfo[i,6]; inc(k); //Internal Cor
        DataToExcel[i-1,k] := pipeinfo[i,7]; inc(k); //External cor
        DataToExcel[i-1,k] := pipeinfo[i,8]; inc(k); //tension Adj
        DataToExcel[i-1,k] := Pinfo[i,10]; inc(k);
        { Data eacho of external and internal fluid properties }
        DataToExcel[i-1,k] := Extfluid[i,1]; inc(k); // External FLuid Elevation
        DataToExcel[i-1,k] := Extfluid[i,2]; inc(k); // External Fluid Density
        DataToExcel[i-1,k] := Extfluid[i,3]; inc(k); // External Fluid Pressure at reference elevation
        if IntfluidElementOverride = False then
          begin;
          FluidVal := eleminfo[i,4];
          DataToExcel[i-1,k] := Fluid[FluidVal,1]; inc(k); // Internal Fluid Elevation
          DataToExcel[i-1,k] := Fluid[FluidVal,2]; inc(k); // Internal Fluid Density
          DataToExcel[i-1,k] := Fluid[FluidVal,3]; inc(k); // Internal Fluid Pressure at reference elevation
          end
        else
          begin;
          DataToExcel[i-1,k] := IntFluidElement[i,1]; inc(k); // Internal Fluid Elevation
          DataToExcel[i-1,k] := IntFluidElement[i,2]; inc(k); // Internal Fluid Density
          DataToExcel[i-1,k] := IntFluidElement[i,3]; inc(k); // Internal Fluid Pressure at reference elevation
          end;
        if DNVCodeCheck then
         begin;
          for j := 1 to 7 do begin; DataToExcel[i-1,k] := DNVFactors[i,j]; inc(k); end;
          for j := 1 to 6 do begin; DataToExcel[i-1,k] := DNVMatSpec[i,j]; inc(k); end;
          for j := 1 to 2 do begin; DataToExcel[i-1,k] := DNVMults[i,j]; inc(k); end;
         end;
        if DNVPInput then
         begin;
          for j := 1 to 6 do begin; DataToExcel[i-1,k] := DNVPres[i,j]; inc(k); end;
         end;
        if ISOCodeCheck then
         begin;
         DataToExcel[i-1,k] := pipeinfo[i,10]; inc(k);
         DataToExcel[i-1,k] := pipeinfo[i,11]; inc(k);
         DataToExcel[i-1,k] := pipeinfo[i,12]; inc(k);
         DataToExcel[i-1,k] := pipeinfo[i,13]; inc(k);
         end;
        if APISTD2RDCodeCheck = true then
        begin;
         if Collapse_Type[i] = true then
           DataToExcel[i-1,k] := 'Plastic'
           else
           DataToExcel[i-1,k] := 'Elastic';
         inc(k); // Collapse Equation
         DataToExcel[i-1,k] := SAFactor[i]; inc(k); // Strain Amplification Factor
         DataToExcel[i-1,k] := BurstK[i]; inc(k); // Mechanical Variability
         DataToExcel[i-1,k] := FcM4[i]; inc(k); // Collapse Factor fc
        if  globalfactor = false then
          begin;
          DataToExcel[i-1,k] := Temperature[i,1]; inc(k); // Temperature
          DataToExcel[i-1,k] := Round(Temperature[i,2]); inc(k) // Thermal Table
          end
          else
          begin;
          DataToExcel[i-1,k] := TemperatureFactor[i]; inc(k); // Temperature Factor Yield
          DataToExcel[i-1,k] := TemperatureFactorTensile[i]; inc(k);
          end;
        end;
        end;
       countlab := 0;
       if TotalLabelE > 0 then
        begin;
        setlength(DataSToExcel,TotalLabelE,4);
        setlength(Data2ToExcel,TotalLabelE,51);
        for j := 1 to TotalLabelE do
           begin;
           previouslab := 0; different := false;
           for i := 1 to Totalelements do
              begin;
              if (eleminfo[i,1] >= LabelElemI[j,1]) and (eleminfo[i,1] <= LabelElemI[j,2]) then
                begin;
                if previouslab = 0 then previouslab := i
                else
                 begin; // compare data array to see if different.
                 for k2 := 3 to k do
                  if DataToExcel[i-1,k2] <> DataToExcel[previouslab-1,k2] then different := true;
                 end;
                end;
              end;
           if (different = false) and (previouslab <> 0) and (pos('_First',LabelElem[j]) = 0) and (pos('_Last',LabelElem[j]) = 0) then
           begin;
           DataSToExcel[countlab,0] := '{'+LabelElem[j]+'}';
           DataSToExcel[countlab,1] := 'N/A';
           DataSToExcel[countlab,2] := 'N/A';
           for k2 := 3 to k do Data2ToExcel[countlab,k2-3] := DataToExcel[previouslab-1,k2];
           inc(countlab);
           end;
           end;
       TransferArr  := Data2ToExcel;
       Sheet.Range[Sheet.Cells[9,4],Sheet.Cells[9+countlab-1,1+k]] := TransferArr;
       Finalize(Data2ToExcel); Finalize(TransferArr);
       TransferArr2 := DataSToExcel;
       Sheet.Range[Sheet.Cells[9,1],Sheet.Cells[9+countlab-1,3]] := TransferArr2;
       Finalize(DataSToExcel); Finalize(TransferArr2);
       end;
       TransferArr := DataToExcel;
       Sheet.Range[Sheet.Cells[9+countlab,1],Sheet.Cells[9+Countlab+TotalElements-1,1+k]] := TransferArr;
       Finalize(DataToExcel);  Finalize(TransferArr);
       end;
      if LowerCase(Tempstr) = 'forces+stresses' then
       begin;
       CFN := 1;
       setlength(HeaderToExcel,10,40);
       HeaderToExcel[0,1] := 'Created by : ' + PROG_NAME + '       by 2H Offshore Engineering';
       HeaderToExcel[1,1] := 'Version    : ' + Version;
       HeaderToExcel[2,1] := 'Title      : ' + Title;
       HeaderToExcel[3,1] := 'Filename   : ' + filename;
       HeaderToExcel[3,5] := '.INF File  : ' + infFile;
       HeaderToExcel[4,5] := 'FORCES AND STRESSES';
       k := 0;
       AddHeader('BLANK','Element','Number',5,k); inc(k);
       AddHeader('BLANK','Local','Node',5,k); inc(k);
       AddHeader('Effective Tension','Min','kN',5,k); inc(k);
       AddHeader('BLANK','Max','kN',5,k); inc(k);
       AddHeader('Pressure','Internal','MPa',5,k); inc(k);
       Addheader('BLANK','External','MPa',5,k); inc(k);
       AddHeader('Shear Y','Min','kN',5,k); inc(k);
       AddHeader('BLANK','Max','kN',5,k); inc(k);
       AddHeader('Shear Z','Min','kN',5,k); inc(k);
       AddHeader('BLANK','Max','kN',5,k); inc(k);
       AddHeader('Torque','Min','kNm',5,k); inc(k);
       AddHeader('BLANK','Max','kNm',5,k); inc(k);
       AddHeader('Bending Moment Y','Min','kNm',5,k); inc(k);
       AddHeader('BLANK','Max','kNm',5,k); inc(k);
       AddHeader('Bending Moment Z','Min','kNm',5,k); inc(k);
       AddHeader('BLANK','Max','kNm',5,k); inc(k);
       AddHeader('Resultant Bending Moment','Min','kNm',5,k); inc(k);
       AddHeader('BLANK','Max','kNm',5,k); inc(k);
       AddHeader('Bending Radius','Min','m',5,k); inc(k);
       AddHeader('BLANK','Max','m',5,k); inc(k);
       if ISOCodeCheck then
        begin;
        Addheader('ISO Design','Factor','0.67',5,k); inc(k);
        AddHeader('ISO Design','Factor','0.80',5,k); inc(k);
        AddHeader('ISO Design','Factor','0.90',5,k); inc(k);
        AddHeader('ISO Design','Factor','1.00',5,k); inc(k);
        end
        else
        begin;
        AddHeader('Von','Mises','/Yield',5,k);inc(k);
        end;
       if DNVCodeCheck then
        begin;
        AddHeader('DNV Code Check','Max','Unity Check Value',5,k); inc(k);
        AddHeader('BLANK','Internal','BLANK',5,k); inc(k);
        AddHeader('BLANK','External','BLANK',5,k); inc(k);
        // To add DNV square root and utlisation //
        AddHeader('DNV Code Check Alternative Scaled Values','Max','Unity Value',5,k); inc(k);
        AddHeader('BLANK','Internal','BLANK',5,k); inc(k);
        AddHeader('BLANK','External','BLANK',5,k); inc(k);
        end;
       if APISTD2RDCodeCheck = true then
        begin;
        AddHeader('API-STD-2RD Code Check','Method 1','Utilisation',5,k); inc(k);
        AddHeader('BLANK','Method 2','Utilisation',5,k); inc(k);
        AddHeader('BLANK','Method 3','Util/Internal',5,k); inc(k);
        AddHeader('BLANK','Method 3','Util/External',5,k); inc(k);
        AddHeader('BLANK','Method 4','Load Utilisation',5,k); inc(k);
        AddHeader('BLANK','Method 4','Strain/Utilisation',5,k); inc(k);
        AddHeader('Bending Moment','/0.9 Bending','Capacity',5,k); inc(k);
        end;
       TransferArr2 := HeaderToExcel;
       Sheet.Range[Sheet.Cells[1,1],Sheet.Cells[9,1+k]] := TransferArr2;
       Finalize(HeaderToExcel); Finalize(TransferArr2);
       { output data forces and stress... }
       countrow := -1;
       sprCount := 1;
       flexjntCount := 0;
       setlength(DataToExcel,TotalElements,40);
       for i := 1 to totalElements do
        begin;
        elem := elemInfo[i,1];
        if (elem = spring[sprCount]) or (elem = flexjoint[FlexJntCount,0]) then
         begin;
         if elem = spring[sprCount] then inc(sprCount)
         else if flexJntCount < High(flexJoint) then inc(FlexJntCount);
         end
        else
         begin;
            // determine which internode to output
            j := 1; k := 0; inc(countrow);
            vonm1 := StressCheckA[i,1,1];
            vonm2 := StressCheckA[i,2,1];
            if vonm2 > vonm1 then j := 2;
            if elemSwitch[i] = 1 then j := 1;
            if elemSwitch[i] = 2 then j := 2;
            DataToExcel[countrow,k] := elem; inc(k);
            DataToExcel[countrow,k] := j; inc(k);
            for k2 := 1 to 16 do
             begin;
             if k2 < 15 then begin; DataToExcel[countrow,k] := loads[i,k2,j,CFN]/1000; inc(k); end;
             if k2 >= 15 then begin; DataToExcel[countrow,k] := loads[i,k2,j,CFN]; inc(k); end;
             if k2 = 2 then
                begin;
                DataToExcel[countrow,k] := Pinfo[i,2*j]; inc(k);
                DataToExcel[countrow,k] := Pinfo[i,2*j+1]; inc(k);
                end;
             end;
            if ISOCodeCheck then
              begin;
              DataToExcel[countrow,k] := StressCheckA[i,j,2]; inc(k);
              DataToExcel[countrow,k] := StressCheckA[i,j,3]; inc(k);
              DataToExcel[countrow,k] := StressCheckA[i,j,4]; inc(k);
              DataToExcel[countrow,k] := StressCheckA[i,j,5]; inc(k);
              end
              else
              begin;
              DataToExcel[countrow,k] := StressCheckA[i,j,1]; inc(k);
              end;
            if DNVCodeCheck then
              begin;
              DataToExcel[countrow,k] := StressCheckA[i,j,6]; inc(k);
              if StressCheckA[i,j,7] < 0 then
              DataToExcel[countrow,k] := 'N/A'
              else
              DataToExcel[countrow,k] := StressCheckA[i,j,7]; inc(k);
              if StressCheckA[i,j,8] < 0 then
              DataToExcel[countrow,k] := 'N/A'
              else
              DataToExcel[countrow,k] := StressCheckA[i,j,8]; inc(k);
              DataToExcel[countrow,k] := StressCheckA[i,j,9]; inc(k);
              if StressCheckA[i,j,10] < 0 then
              DataToExcel[countrow,k] := 'N/A'
              else
              DataToExcel[countrow,k] := StressCheckA[i,j,10]; inc(k);
              if StressCheckA[i,j,11] < 0 then
              DataToExcel[countrow,k] := 'N/A'
              else
              DataToExcel[countrow,k] := StressCheckA[i,j,11]; inc(k);
              end;
            if APISTD2RDCodeCheck = true then
              begin;
              DNVutlInt := max(DNVInt[i,1], DNVInt[i,2]);
              DNVutlExt := max(DNVExt[i,1], DNVExt[i,2]);
              if DNVutlExt >= 0 then DNVUtlExtSR := sqrt(DNVutlExt)
              else
              DNVUtlExtSR := -1;
              DataToExcel[countrow,k] := STD2RDOut[i,2]; inc(k);  // Method 1 Util
              DataToExcel[countrow,k] := STD2RDOut[i,3]; inc(k);  // Method 2 Util
              if DNVUtlInt < 0 then
              DataToExcel[countrow,k] := 'N/A'
              else
              DataToExcel[countrow,k] := DNVUtlInt; inc(k);  // Method 3 Int
              if DNVUtlExtSR < 0 then
              DataToExcel[countrow,k] := 'N/A'
              else
              DataToExcel[countrow,k] := DNVUtlExtSR; inc(k);  // Method 3 Ext
              if APISTDValMethod4[i] = true then
              begin;
              DataToExcel[countrow,k] := STD2RDOut[i,4]; inc(k); // method 4 Load Util
              DataToExcel[countrow,k] := STD2RDOut[i,5]; inc(k);  // method 4 Strain Util
              end
              else
              begin;
              DataToExcel[countrow,k] := 'Insufficient Data'; inc(k); // method 4 Load Util
              DataToExcel[countrow,k] := 'Insufficient Data'; inc(k);  // method 4 Strain Util
              end;
              if STD2RDOut[i,1] < 1E8 then
              begin; DataToExcel[countrow,k] := STD2RDOut[i,1]; inc(k); end  // bm /0.9 MMax
              else
              begin; DataToExcel[countrow,k] := 'Fail Mmax=0'; inc(k); end;
              end;
         end;
        end;
        countlab := 0;
       if TotalLabelE > 0 then
        begin;
        setlength(DataSToExcel,TotalLabelE+1,4);
        setlength(Data2ToExcel,TotalLabelE+1,40);
        for j := 1 to TotalLabelE do
           begin;
           endlab := 0; firstlab := 0; different := false;
           SprCount := 1;
           FlexjntCount := 0;
           for i := 1 to Totalelements do
              begin;
               if (eleminfo[i,1] = spring[sprCount]) or (eleminfo[i,1] = flexjoint[FlexJntCount,0]) then
                  begin;
                  if elem = spring[sprCount] then inc(sprCount)
                  else if flexJntCount < High(flexJoint) then inc(FlexJntCount);
                  end
              else if (eleminfo[i,1] >= LabelElemI[j,1]) and (eleminfo[i,1] <= LabelElemI[j,2]) then
                begin;
                if firstlab = 0 then firstlab := i
                else
                 begin; // compare data array to see if different.
                 endlab := i;
                 end;
                end;
              end;
            if endlab = 0 then endlab := firstlab;
           if (firstlab <> 0) and (pos('_First',LabelElem[j]) = 0) and (pos('_Last',LabelElem[j]) = 0) then
           begin;
           DataSToExcel[countlab,0] := '{'+LabelElem[j]+'}';
           DataSToExcel[countlab,1] := 'N/A';
           for i := firstlab to endlab do
           for j2 := 0 to countrow do
           if eleminfo[i,1] = DataToExcel[j2,0] then
           begin;
           for k2 := 2 to (k-1) do
            begin;
               if (k2 = 2) or (k2=6) or (k2=8) or (k2=10) or (k2=12) or (k2=14) or (k2=16) or (k2=18) then
                 begin;
                 if i = firstlab then Data2ToExcel[countlab,k2-2] := DataToExcel[j2,k2];
                 tempstring1 := Data2ToExcel[countlab,k2-2];
                 tempstring2 := DataToExcel[j2,k2];
                 if (tempstring1 = 'N/A') or (tempstring2 = 'N/A') then
                   Data2ToExcel[countlab,k2-2] := 'N/A'
                 else
                 if (tempstring1 ='Insufficient Data') or (tempstring2 = 'Insufficient Data') then
                   Data2ToExcel[countlab,k2-2] := 'Insufficient Data'
                 else
                 if (tempstring1 ='Fail Mmax=0') or (tempstring2 = 'Fail Mmax=0') then
                   Data2ToExcel[Countlab,k2-2] := 'Fail Mmax=0'
                 else
                 begin;
                  tempmin1 := Data2ToExcel[countlab,k2-2];
                  tempmin2 := DataToExcel[j2,k2];
                  tempmin := min(tempmin1,tempmin2);
                  Data2ToExcel[countlab,k2-2] := tempmin;
                 end;
                 end
               else if (k2=4) or (k2=5) then
                 begin;
                 // do nothing.
                 end
               else
                 begin;
                 if i = firstlab then Data2ToExcel[countlab,k2-2] := DataToExcel[j2,k2];
                 tempstring1 := Data2ToExcel[countlab,k2-2];
                 tempstring2 := DataToExcel[j2,k2];
                 if (tempstring1 ='Insufficient Data') or (tempstring2 = 'Insufficient Data') then
                   Data2ToExcel[Countlab,k2-2] := 'Insufficient Data'
                  else
                  if (tempstring1 ='Fail Mmax=0') or (tempstring2 = 'Fail Mmax=0') then
                   Data2ToExcel[Countlab,k2-2] := 'Fail Mmax=0'
                  else
                  if (tempstring1 = 'N/A') or (tempstring2 = 'N/A') then
                    Data2ToExcel[countlab,k2-2] := DataToExcel[j2,k2]
                  else
                  begin;
                    tempmax1 := Data2ToExcel[countlab,k2-2];
                    tempmax2 := DataToExcel[j2,k2];
                    tempmax := max(tempmax1,tempmax2);
                    Data2ToExcel[countlab,k2-2] := tempmax;
                  end;
                 end;
             end;
           end;
          inc(countlab)
          end;
       end;
       TransferArr  := Data2ToExcel;
       Sheet.Range[Sheet.Cells[9,3],Sheet.Cells[9+countlab-1,1+k]] := TransferArr;
       Finalize(Data2ToExcel); Finalize(TransferArr);
       TransferArr2 := DataSToExcel;
       Sheet.Range[Sheet.Cells[9,1],Sheet.Cells[9+countlab-1,2]] := TransferArr2;
       Finalize(DataSToExcel); Finalize(TransferArr2);
       end;
       TransferArr := DataToExcel;
       Sheet.Range[Sheet.Cells[9+countlab,1],Sheet.Cells[9+countlab+Countrow,1+k]] := TransferArr;
       Finalize(DataToExcel); Finalize(TransferArr);
       end;
      if LowerCase(Tempstr) = 'displacements' then
       begin;
       setlength(HeaderToExcel,10,40);
       HeaderToExcel[0,1] := 'Created by : ' + PROG_NAME + '       by 2H Offshore Engineering';
       HeaderToExcel[1,1] := 'Version    : ' + Version;
       HeaderToExcel[2,1] := 'Title      : ' + Title;
       HeaderToExcel[3,1] := 'Filename   : ' + filename;
       HeaderToExcel[3,5] := '.INF File  : ' + infFile;
       HeaderToExcel[4,5] := 'DISPLACEMENTS';
       k := 0;
       AddHeader('Node','BLANK','BLANK',5,k); inc(k);
       AddHeader('LOCATION','X','m',5,k); inc(k);
       AddHeader('BLANK','Y','m',5,k); inc(k);
       AddHeader('BLANK','Z','m',5,k); inc(k);
       AddHeader('VERTICAL DISPLACEMENT','Min','m',5,k); inc(k);
       AddHeader('BLANK','Max','m',5,k); inc(k);
       AddHeader('HORIZONTAL Y DISPLACEMENT','Min','m',5,k); inc(k);
       AddHeader('BLANK','Max','m',5,k); inc(k);
       AddHeader('HORIZONTAL Z DISPLACEMENT','Min','m',5,k); inc(k);
       AddHeader('BLANK','Max','m',5,k); inc(k);
       AddHeader('ANGLE X','Min','deg',5,k); inc(k);
       AddHeader('BLANK','Max','deg',5,k); inc(k);
       AddHeader('ANGLE Y','Min','deg',5,k); inc(k);
       AddHeader('BLANK','Max','deg',5,k); inc(k);
       AddHeader('ANGLE Z','Min','deg',5,k); inc(k);
       AddHeader('BLANK','Max','deg',5,k); inc(k);
       { Output to Excel }
      TransferArr2 := HeaderToExcel;
       Sheet.Range[Sheet.Cells[1,1],Sheet.Cells[7,1+k]] := TransferArr2;
       Finalize(HeaderToExcel); Finalize(transferArr2);
       { displacement data }
      setlength(DataToExcel,TotalNodes+1,20);
       for i := 1 to TotalNodes do
       begin;
        k := 0;
        DataToExcel[i-1,k] := nodeinfoI[i]; inc(k); //Node
        for j:= 1 to 15 do
         begin;
         DataToExcel[i-1,k] := nodeinfo[i,j]; inc(k);
         end;
       end;
       countlab := 0;
       { If nodal label at length one set to reoutput }
        if TotalLabelN > 0 then
        begin;
        setlength(DataSToExcel,TotalLabelN+1,1);
        setlength(Data2ToExcel,TotalLabelN+1,40);
        for i := 1 to TotalNodes do
           begin;
           for j := 1 to TotalLabelN do
             if (NodeInfoI[i] = LabelNodeI[j,1]) and (LabelNodeI[j,1] = LabelNodeI[j,2]) then
              begin;
              DataSToExcel[countlab,0] := '{'+LabelNode[j]+'}';
              for j2 := 1 to 15 do
              Data2ToExcel[countlab,j2-1] := nodeinfo[i,j2];
              inc(countlab);
              end;
           end;
        if countlab > 0 then
        begin;
        TransferArr  := Data2ToExcel;
        Sheet.Range[Sheet.Cells[8,2],Sheet.Cells[8+countlab-1,16]] := TransferArr;
        Finalize(Data2ToExcel); Finalize(TransferArr);
        TransferArr2 := DataSToExcel;
        Sheet.Range[Sheet.Cells[8,1],Sheet.Cells[8+countlab-1,1]] := TransferArr2;
        Finalize(DataSToExcel); Finalize(TransferArr2);
        end;
        end;
       { end of Nodal labels for Displacement}
       TransferArr := DataToExcel;
       Sheet.Range[Sheet.Cells[8+countlab,1],Sheet.Cells[8+countlab+TotalNodes-1,16]] := TransferArr;
       Finalize(DataToExcel); Finalize(TransferArr);
       end;
      if LowerCase(Tempstr) = 'springs' then
       begin;
       setlength(HeaderToExcel,11,40);
       HeaderToExcel[0,0] := 'Created by : ' + PROG_NAME + '       by 2H Offshore Engineering';
       HeaderToExcel[1,0] := 'Version    : ' + Version;
       HeaderToExcel[2,0] := 'Title      : ' + Title;
       HeaderToExcel[3,0] := 'Filename   : ' + filename;
       HeaderToExcel[3,4] := '.INF File  : ' + infFile;
       HeaderToExcel[4,4] := ' Spring Data';
       k := 0;
       AddHeader('ELEMENT','BLANK','BLANK',5,k); inc(k);
       AddHeader('Spring Force','Min','N',5,k); inc(k);
       AddHeader('BLANK','Max','N',5,k); inc(k);
       AddHeader('Spring Stroke','Min','m',5,k); inc(k);
       AddHeader('BLANK','Max','m',5,k); inc(k);
       if totalsprings = 0 then HeaderToExcel[9,2] := 'NO SPRINGS IN MODEL';
       TransferArr2 := HeaderToExcel;
       Sheet.Range[Sheet.Cells[1,1],Sheet.Cells[10,1+k]] := TransferArr2;
       Finalize(HeaderToExcel); Finalize(TransferArr2);
       { Add in spring data }
      if totalsprings > 0 then
       begin;
       setlength(DataToExcel,totalsprings+2,10);
       for i := 1 to totalsprings do
        begin;
        DataToExcel[i,0] := Spring[i];
        DataToExcel[i,1] := SpringData[i,5];
        DataToExcel[i,2] := SpringData[i,4];
        DataToExcel[i,3] := SpringData[i,7];
        DataToExcel[i,4] := SpringData[i,6];
        end;
       countlab := 0;
       if TotalLabelE > 0 then
        begin;
        setlength(DataSToExcel,TotalLabelE+1,4);
        setlength(Data2ToExcel,TotalLabelE+1,40);
        countlab := 0;
        for i := 1 to TotalSprings do
        for j := 1 to TotalLabelE do
           begin;
           if LabelElemI[j,1] = Spring[i] then
             if LabelElemI[j,2] = Spring[i] then
               begin;
               DataSToExcel[countlab,0] := '{' + LabelElem[j] + '}';
               Data2ToExcel[countlab,0] := SpringData[i,5];
               Data2ToExcel[countlab,1] := SpringData[i,4];
               Data2ToExcel[countlab,2] := SpringData[i,7];
               Data2ToExcel[countlab,3] := SpringData[i,6];
               inc(Countlab);
               end;
           end;
       TransferArr2 := DataSToExcel;
       Sheet.Range[Sheet.Cells[9,1],Sheet.Cells[9+countlab-1,1]] := TransferArr2;
       Finalize(DataSToExcel); Finalize(TransferArr2);
       TransferArr := Data2ToExcel;
       Sheet.Range[Sheet.Cells[9,2],Sheet.Cells[9+countlab-1,10]] := TransferArr;
       Finalize(Data2ToExcel); Finalize(TransferArr);
       end;
       TransferArr := DataToExcel;
       Sheet.Range[Sheet.Cells[9+countlab,1],Sheet.Cells[9+totalsprings+countlab,10]] := TransferArr;
       Finalize(DataToExcel); Finalize(TransferArr);
       end;
       end;
      if LowerCase(Tempstr) = 'flexjoints' then
       begin;
       setlength(HeaderToExcel,11,40);
       HeaderToExcel[0,1] := 'Created by : ' + PROG_NAME + '       by 2H Offshore Engineering';
       HeaderToExcel[1,1] := 'Version    : ' + Version;
       HeaderToExcel[2,1] := 'Title      : ' + Title;
       HeaderToExcel[3,1] := 'Filename   : ' + filename;
       HeaderToExcel[3,5] := '.INF File  : ' + infFile;
       HeaderToExcel[4,5] := ' Flex Joint Data and Articulation';
       HeaderToExcel[5,2] := 'Note : Angle Timetrace method is valid for ALL flex joint stiffness';
       HeaderToExcel[6,2] := '     : Bending Moment method is not valid for LOW flex joint stiffness';
       k := 0;
       AddHeader('Element','Number','BLANK',7,k); inc(k);
       Addheader('Flex Joint','Type','BLANK',7,k); inc(k);
       Addheader('Method','BLANK','BLANK',7,k); inc(k);
       Addheader('Bending Moment','Min','kNm',7,k); inc(k);
       Addheader('BLANK','Max','kNm',7,k); inc(k);
       Addheader('ANGLE','Min','deg',7,k); inc(k);
       Addheader('BLANK','Max','deg',7,k); inc(k);
       Addheader('ANGLE TAKEN BETWEEN','ELEMENT','BLANK',7,k); inc(k);
       Addheader('BLANK','ELEMENT','BLANK',7,k); inc(k);
       if TotalFlexjoints = 0 then HeaderToExcel[10,2] := ' NO FLEX JOINTS IN MODEL';
       TransferArr2 := HeaderToExcel;
       Sheet.Range[Sheet.Cells[1,1],Sheet.Cells[11,1+k]] := TransferArr2;
       Finalize(HeaderToExcel); Finalize(TransferArr2);
       { output flexjoint info }
      if TotalFlexjoints > 0 then
       begin;
       Setlength(DataSToExcel,High(outputFJarr)+2,12);
       for i := 0 to High(outputFJarr) do
        begin;
        for j := 0 to 8 do
         begin;
         DataSToExcel[i,j] := outputFJArr[i,j];
         { flip max/min to min/max }
         if j = 3 then DataSToExcel[i,j] := outputFJArr[i,j+1];
         if j = 4 then DataSToExcel[i,j] := outputFJArr[i,j-1];
         if j = 5 then DataSToExcel[i,j] := outputFJArr[i,j+1];
         if j = 6 then DataSToExcel[i,j] := outputFJArr[i,j-1];
         end;
        end;
        countlab := 0;
      if TotalLabelE > 0 then
       begin;
        setlength(DataS2ToExcel,TotalLabelE+1,10);
        countlab := 0;
        for i := 0 to High(outputFJArr) do
        for j := 1 to TotalLabelE do
           begin;
           if LabelElemI[j,1] = flexjoint[i,0] then
             if LabelElemI[j,2] = flexjoint[i,0] then
               if (Pos('_Last',LabelElem[j]) = 0) and (Pos('_First',LabelElem[j])= 0) then
               begin;
               DataS2ToExcel[countlab,0] := '{' + LabelElem[j] + '}';
               for k := 1 to 8 do
               DataS2ToExcel[countlab,k] := DataSToExcel[i,k];
               inc(Countlab);
               end;
           end;
       TransferArr2 := DataS2ToExcel;
       Sheet.Range[Sheet.Cells[11,1],Sheet.Cells[11+countlab-1,10]] := TransferArr2;
       Finalize(DataS2ToExcel); Finalize(TransferArr2);
       end;
       TransferArr := DataSToExcel;
       Sheet.Range[Sheet.Cells[11+countlab,1],Sheet.Cells[11+High(outputFJarr)+countlab,10]] := TransferArr;
       Finalize(DataSToExcel);  Finalize(TransferArr);
       end;
       end;
     end;
     for i := 1 to numWorkSheets do
          begin;
          Sheet := Excel.WorkSheets[i];
          Tempstr := Sheet.Name;
          if LowerCase(Tempstr) = 'cover' then begin; Sheet.Cells[4,4].Value := 'First'; end;
          end;
     Excel.DisplayAlerts := False;
     if FileExists(outputfilename) = True then
     if DeleteFile(outputfilename) = False then
        begin;
        writeln('Unable to delete ',outputfilename);
        writeln('No data saved');
        OutputToExcel := false;
        end;
     if FileExists(outputfilename) = False then
     begin;
     Excel.Application.Workbooks[1].SaveAs(OutputFileName,52);
     writeln('Successfully saved Excel File - Stage 1');
     end;

     Excel.Workbooks.Close;
     Excel.IgnoreRemoteRequests := False;
     Excel.Quit;
  end;

  except on E:Exception do
   begin;
   Writeln('Program has encountered error in Excel File Writing');
   Writeln('Error Message ' +E.Message+'.');
   Excel.DisplayAlerts := False;
   Excel.Workbooks.Close;
   Excel.IgnoreRemoteRequests := False;
   Excel.Quit;
   halt(1);
   end;
  end;

end;


Procedure OutputExcel2; { procedure to output data to excel files... }

var
   Dirlist : string;
   TemplateFileName,TemplateLocation,CurrentDirectory : string;
   OutputFileName : string;
   Sheet, Excel : OLEVariant;
   TransferArr,TransferArr2 : oleVariant;
   i,ish,j,j2,k, numWorkSheets,countrow,maxlinesTables : integer;
   sprCount,flexjntcount,elem : integer;
   k2,CFN,countlab,previouslab,endlab,firstlab : integer;
   vonm1,vonm2,tempmin,tempmax,tempmin1,tempmin2,tempmax1,tempmax2 : real;
   different,staticdel,functionaldel : boolean;
   outfileno,staticno,functionalno : integer;
   TempStr : string;

begin;

   Staticdel := false; Functionaldel := false;
   CurrentDirectory := GetCurrentDir;
   if INFnameInOutput = false then
      OutputFileName := CurrentDirectory + '\' + filename + '_2HSTRE3D.xlsm'
   else
      OutputFileName := CurrentDirectory + '\' + filename + '_' + infFile + '_2HSTRE3D.xlsm';
   if FileExists(OutputFileName) = false then exit;

  try
     Excel := CreateOleObject('Excel.Application');

  if VarIsEmpty(Excel) = false then
     begin;
     Excel.Visible := false; ///false;
     Excel.IgnoreRemoteRequests := True;
     Excel.Workbooks.Open(OutputFileName, ReadOnly := False);
     { determine number of worksheets }
     numWorkSheets := Excel.Worksheets.Count;
    for ish := 1 to numWorkSheets do
      begin;
      Sheet := Excel.Worksheets[ish];
      Tempstr := Sheet.Name;
      if (LowerCase(tempstr) = 'functional loads') or (Lowercase(tempstr) = 'static loads') then
      begin;
      OutfileNo := 4;
      if (LowerCase(tempstr) = 'functional loads') then OutFileNo := 2;
      if (LowerCase(tempstr) = 'static loads') then OutFileNo := 3;
      if (MaxFileNo < OutfileNo) then  // delete the worksheet.
        begin;
         if OutfileNo = 2 then begin; FunctionalDel := true; end;
         if OutfileNo = 3 then begin; StaticDel := true;  end;
        end
      else
      begin;
       setlength(HeaderToExcel,10,40);
       HeaderToExcel[0,1] := 'Created by : ' + PROG_NAME + '       by 2H Offshore Engineering';
       HeaderToExcel[1,1] := 'Version    : ' + Version;
       HeaderToExcel[2,1] := 'Title      : ' + Title;
       HeaderToExcel[3,1] := 'Filename   : ' + filename;
       HeaderToExcel[3,5] := '.INF File  : ' + infFile;
       if OutfileNo = 2 then HeaderToExcel[4,6] := 'FORCES Data Echo of Functional Loads';
       if OutfileNo = 3 then HeaderToExcel[4,6] := 'FORCES Data Echo of Static Loads';
       k := 0;
       AddHeader('BLANK','Element','Number',5,k); inc(k);
       AddHeader('BLANK','Local','Node',5,k); inc(k);
       AddHeader('Effective Tension','Min','kN',5,k); inc(k);
       AddHeader('BLANK','Max','kN',5,k); inc(k);
       AddHeader('Shear Y','Min','kN',5,k); inc(k);
       AddHeader('BLANK','Max','kN',5,k); inc(k);
       AddHeader('Shear Z','Min','kN',5,k); inc(k);
       AddHeader('BLANK','Max','kN',5,k); inc(k);
       AddHeader('Torque','Min','kNm',5,k); inc(k);
       AddHeader('BLANK','Max','kNm',5,k); inc(k);
       AddHeader('Bending Moment Y','Min','kNm',5,k); inc(k);
       AddHeader('BLANK','Max','kNm',5,k); inc(k);
       AddHeader('Bending Moment Z','Min','kNm',5,k); inc(k);
       AddHeader('BLANK','Max','kNm',5,k); inc(k);
       AddHeader('Resultant Bending Moment','Min','kNm',5,k); inc(k);
       AddHeader('BLANK','Max','kNm',5,k); inc(k);
       AddHeader('Bending Radius','Min','m',5,k); inc(k);
       AddHeader('BLANK','Max','m',5,k); inc(k);

       TransferArr2 := HeaderToExcel;
       Sheet.Range[Sheet.Cells[1,1],Sheet.Cells[8,40]] := TransferArr2;
       Finalize(HeaderToExcel); Finalize(TransferArr2);
        countrow := -1;
       sprCount := 1;
       flexjntCount := 0;
       setlength(DataToExcel,TotalElements,40);
       for i := 1 to totalElements do
        begin;
        elem := elemInfo[i,1];
        if (elem = spring[sprCount]) or (elem = flexjoint[FlexJntCount,0]) then
         begin;
         if elem = spring[sprCount] then inc(sprCount)
         else if flexJntCount < High(flexJoint) then inc(FlexJntCount);
         end
        else
         begin;
            // determine which internode to output
            j := 1; k := 0; inc(countrow);
            CFN := OutFileNo;
            vonm1 := StressCheckA[i,1,1];
            vonm2 := StressCheckA[i,2,1];
            if vonm2 > vonm1 then j := 2;
            if elemSwitch[i] = 1 then j := 1;
            if elemSwitch[i] = 2 then j := 2;
            DataToExcel[countrow,k] := elem; inc(k);
            DataToExcel[countrow,k] := j; inc(k);
            for k2 := 1 to 16 do
             begin;
             if k2 < 15 then begin; DataToExcel[countrow,k] := loads[i,k2,j,CFN]/1000; inc(k); end;
             if k2 >= 15 then begin; DataToExcel[countrow,k] := loads[i,k2,j,CFN]; inc(k); end;
             end;
         end;
        end;
        countlab := 0;
       if TotalLabelE > 0 then
        begin;
        setlength(DataSToExcel,TotalLabelE+1,4);
        setlength(Data2ToExcel,TotalLabelE+1,40);
        for j := 1 to TotalLabelE do
           begin;
           endlab := 0; firstlab := 0; different := false;
           SprCount := 1;
           FlexjntCount := 0;
           for i := 1 to Totalelements do
              begin;
               if (eleminfo[i,1] = spring[sprCount]) or (eleminfo[i,1] = flexjoint[FlexJntCount,0]) then
                  begin;
                  if elem = spring[sprCount] then inc(sprCount)
                  else if flexJntCount < High(flexJoint) then inc(FlexJntCount);
                  end
              else if (eleminfo[i,1] >= LabelElemI[j,1]) and (eleminfo[i,1] <= LabelElemI[j,2]) then
                begin;
                if firstlab = 0 then firstlab := i
                else
                 begin; // compare data array to see if different.
                 endlab := i;
                 end;
                end;
              end;
            if endlab = 0 then endlab := firstlab;
           if (firstlab <> 0) and (pos('_First',LabelElem[j]) = 0) and (pos('_Last',LabelElem[j]) = 0) then
           begin;
           DataSToExcel[countlab,0] := '{'+LabelElem[j]+'}';
           DataSToExcel[countlab,1] := 'N/A';
           for i := firstlab to endlab do
           for j2 := 0 to countrow do
           if eleminfo[i,1] = DataToExcel[j2,0] then
           begin;
           for k2 := 2 to (k-1) do
            begin;
               if (k2 = 2) or (k2=4) or (k2=6) or (k2=8) or (k2=10) or (k2=12) or (k2=14) or (k2=16) then
                 begin;
                 if i = firstlab then Data2ToExcel[countlab,k2-2] := DataToExcel[j2,k2];
                 tempmin1 := Data2ToExcel[countlab,k2-2];
                 tempmin2 := DataToExcel[j2,k2];
                 tempmin := min(tempmin1,tempmin2);
                 Data2ToExcel[countlab,k2-2] := tempmin;
                 end
               else
                 begin;
                 if i = firstlab then Data2ToExcel[countlab,k2-2] := DataToExcel[j2,k2];
                 tempmax1 := Data2ToExcel[countlab,k2-2];
                 tempmax2 := DataToExcel[j2,k2];
                 tempmax := max(tempmax1,tempmax2);
                 Data2ToExcel[countlab,k2-2] := tempmax;
                 end;
             end;
           end;
          inc(countlab)
          end;
       end;
       TransferArr  := Data2ToExcel;
       Sheet.Range[Sheet.Cells[9,3],Sheet.Cells[9+countlab-1,1+k]] := TransferArr;
       Finalize(Data2ToExcel); Finalize(TransferArr);
       TransferArr2 := DataSToExcel;
       Sheet.Range[Sheet.Cells[9,1],Sheet.Cells[9+countlab-1,2]] := TransferArr2;
       Finalize(DataSToExcel); Finalize(TransferArr2);
       end;
       TransferArr := DataToExcel;
       Sheet.Range[Sheet.Cells[9+countlab,1],Sheet.Cells[9+countlab+Countrow,1+k]] := TransferArr;
       Finalize(DataToExcel); Finalize(TransferArr);
       end;
      end;  // Functional Loads or static loads.
      end;

      if functionaldel = true then
      begin;
      numWorkSheets := Excel.Worksheets.Count;
      for ish := 1 to numWorkSheets do
       begin;
       Sheet := Excel.Worksheets[ish];
       Tempstr := Sheet.Name;
       if (LowerCase(tempstr) = 'functional loads') then
         begin;
         functionalno := ish;
         end;
       end;
      Excel.DisplayAlerts := False;
      Excel.Application.WorkSheets[functionalno].Delete;
      end;
      if staticdel = true then
      begin;
      numWorkSheets := Excel.Worksheets.Count;
      for ish := 1 to numWorkSheets do
       begin;
       Sheet := Excel.Worksheets[ish];
       Tempstr := Sheet.Name;
       if (LowerCase(tempstr) = 'static loads') then
         begin;
         staticno := ish;
         end;
       end;
      Excel.DisplayAlerts := False;
      Excel.Application.WorkSheets[staticno].Delete;
      end;
      numWorkSheets := Excel.Worksheets.Count;
      // end of removal of worksheets.
         for i := 1 to numWorkSheets do
          begin;
          Sheet := Excel.WorkSheets[i];
          Tempstr := Sheet.Name;
          if LowerCase(Tempstr) = 'cover' then begin; Sheet.Cells[4,4].Value := 'Second'; end;
          end;
     Excel.DisplayAlerts := False;
     Excel.Application.Workbooks[1].Save;
     writeln('Successfully saved Excel File - Stage 2');

     Excel.Workbooks.Close;
     Excel.IgnoreRemoteRequests := False;
     Excel.Quit;
  end;

  except on E:Exception do
   begin;
   Writeln('Program has encountered error in Excel File Writing');
   Writeln('Error Message ' +E.Message+'.');
   Excel.DisplayAlerts := False;
   Excel.Workbooks.Close;
   Excel.IgnoreRemoteRequests := False;
   Excel.Quit;
   halt(1);
   end;
  end;

end;


Procedure OutputExcel3;
{ **************************************************************************** }
// Writes data to the following sheets of the output Excel file:
//  - Stress_Breakdown
//  - Warnings
//  - Labels
//  - API-STD-2RD_Tables (optional)
{ **************************************************************************** }
 { procedure to output data to excel files... }
var
  Dirlist : string;
  TemplateFileName,TemplateLocation,CurrentDirectory : string;
  OutputFileName : string;
  Sheet, Excel : OLEVariant;
  TransferArr, TransferArr2 : oleVariant;
  i,ish,j,j2,k, numWorkSheets,countrow,maxlinesTables : integer;
  sprCount,flexjntcount: Integer;
  k2,CFN,countlab,previouslab,endlab,firstlab : integer;
  different : boolean;
  APISTD2RDno : integer;
  vonm1,vonm2,tempmin,tempmax,tempmin1,tempmin2,tempmax1,tempmax2 : real;
  TempStr : string;
  elem: Integer;
  node: Integer;
  elemSetsIdx: Integer;
  lineIdx: Integer;
  elemFirst: Integer;
  elemLast: Integer;
  elemLabel: String;
  lineLabel: String;
  sectLabel: String;
  labelStem: String;
  lastRow: Integer;
  lastCol: Integer;

begin
  CurrentDirectory := GetCurrentDir;

  if INFnameInOutput = False then
    OutputFileName := CurrentDirectory + '\' + filename + '_2HSTRE3D.xlsm'
  else
    OutputFileName := CurrentDirectory + '\' + filename + '_' + infFile + '_2HSTRE3D.xlsm';

  if FileExists(OutputFileName) = False then exit;

  try
    Excel := CreateOleObject('Excel.Application');

    if VarIsEmpty(Excel) = False then
    begin
      Excel.Visible := False;
      Excel.IgnoreRemoteRequests := True;
      Excel.Workbooks.Open(OutputFileName, ReadOnly := False);

      // Determine number of worksheets
      numWorkSheets := Excel.Worksheets.Count;

      for ish := 1 to numWorkSheets do
      begin
        Sheet := Excel.Worksheets[ish];

        //================================
        // Write "Stress_Breakdown" sheet
        //================================
        if LowerCase(Sheet.Name) = 'stress_breakdown' then
        begin
          setlength(HeaderToExcel,10,40);

          HeaderToExcel[0,1] := 'Created by : ' + PROG_NAME + '       by 2H Offshore Engineering';
          HeaderToExcel[1,1] := 'Version    : ' + Version;
          HeaderToExcel[2,1] := 'Title      : ' + Title;
          HeaderToExcel[3,1] := 'Filename   : ' + filename;
          HeaderToExcel[3,5] := '.INF File  : ' + infFile;
          HeaderToExcel[4,5] := ' Detailed Component Stresses Breakdown ';

          k := 0;
          AddHeader('Element','Number','BLANK',5,k); inc(k);
          AddHeader('Axial','Stress','Pa',5,k); inc(k);
          AddHeader('Bending','Stress','Pa',5,k); inc(k);
          AddHeader('Hoop','Stress','Pa',5,k); inc(k);
          AddHeader('Radial','Stress','Pa',5,k); inc(k);
          AddHeader('Shear','Stress','Pa',5,k); inc(k);
          AddHeader('Von Mises','Stress','Pa',5,k); inc(k);

          if STT then
          begin
            Addheader('Effective Tension','Min','kN',5,k); inc(k);
            AddHeader('BLANK','Max','kN',5,k); inc(k);
            AddHeader('Resultant Bending Moment','Min','kNm',5,k); inc(k);
            AddHeader('BLANK','Max','kNm',5,k); inc(k);
          end;

          if APISTD2RDCodeCheck = True then
          begin
            Addheader('API STD 2RD Parameters','TensionMax/Ty','BLANK',5,k); inc(k);
            Addheader('BLANK','MomentMax/My','BLANK',5,k); inc(k);
            Addheader('BLANK','PressureRatio','BLANK',5,k); inc(k);
            Addheader('BLANK','MomentMax/Mp','BLANK',5,k); inc(k);
            Addheader('BLANK','Bending Radius Min','m',5,k); inc(k);
            Addheader('BLANK','Bending Strain Max','BLANK',5,k); inc(k);
          end;

          TransferArr2 := HeaderToExcel;
          Sheet.Range[Sheet.Cells[1,1],Sheet.Cells[8,1+k]] := TransferArr2;
          Finalize(HeaderToExcel); Finalize(TransferArr2);

          // Assemble data
          countrow := 0;
          sprCount := 1;
          flexjntCount := 0;
          SetLength(DataToExcel,TotalElements+10,40);

          for i := 1 to TotalElements do
          begin
            elem := eleminfo[i,1];
            if (elem = spring[sprCount]) or (elem = flexjoint[flexjntCount,0]) then
            begin
              if elem = Spring[sprcount] then
                inc(SprCount)
              else if flexJntCount < High(flexjoint) then
                inc(FlexjntCount);
            end
            else
            begin
              k:= 0;

              // Compare value to write out biggest one
              if Maxvmout[i,2] >= Maxvmout[i,1] then j:= 2 else j := 1;

              DataToExcel[countrow,k] := eleminfo[i,1]; inc(k);
              DataToExcel[countrow,k] := MaxStrax[i,j]; inc(k);
              DataToExcel[countrow,k] := Maxstrbmsign[i,j]; inc(k);
              DataToExcel[countrow,k] := MaxStrhoop[i,j]; inc(k);
              DataToExcel[countrow,k] := MaxStrrad[i,j]; inc(k);
              DataToExcel[countrow,k] := Maxshrcomb[i,j]; inc(k);
              DataToExcel[countrow,k] := Maxvmout[i,j]; inc(k);

              if STT then
              begin
                DataToExcel[countrow,k] := tmin[i]/1000; inc(k);
                DataToExcel[countrow,k] := tmax[i]/1000; inc(k);
                DataToExcel[countrow,k] := mmin[i]/1000; inc(k);
                DataToExcel[countrow,k] := mmax[i]/1000; inc(k);
              end;

              if APISTD2RDCodeCheck = True then
              begin
                DataToExcel[countrow,k] := STD2RDOUT[i,7]; inc(k);
                DataToExcel[countrow,k] := STD2RDOUT[i,8]; inc(k);
                DataToExcel[countrow,k] := STD2RDOUT[i,9]; inc(k);
                DataToExcel[countrow,k] := STD2RDOUT[i,10]; inc(k);
                DataToExcel[countrow,k] := min(loads[i,15,1,1],loads[i,15,2,1]); inc(k);

                if STD2RDOut[i,6] < 0 then
                  DataToExcel[countrow,k] := 'Not Calculated'
                else
                  DataToExcel[countrow,k] := STD2RDOUT[i,6];

                inc(k);
              end;

              inc(countrow);
            end;
          end;

          TransferArr := DataToExcel;
          Sheet.Range[Sheet.Cells[9,1],Sheet.Cells[9+Countrow,1+k]] := TransferArr;
          Finalize(DataToExcel);
          Finalize(TransferArr);
        end;

        //========================
        // Write "Warnings" sheet
        //========================
        if LowerCase(Sheet.Name) = 'warnings' then
        begin
          SetLength(HeaderToExcel,10,40);

          HeaderToExcel[0,0] := 'Created by : ' + PROG_NAME + '       by 2H Offshore Engineering';
          HeaderToExcel[1,0] := 'Version    : ' + Version;
          HeaderToExcel[2,0] := 'Title      : ' + Title;
          HeaderToExcel[3,0] := 'Filename   : ' + filename;
          HeaderToExcel[3,4] := '.INF File  : ' + infFile;
          HeaderToExcel[4,2] := 'Compliation of Warnings';
          k := 5;

          // Output to Excel
          TransferArr2 := HeaderToExcel;
          Sheet.Range[Sheet.Cells[1,1],Sheet.Cells[9,1+k]] := TransferArr2;
          Finalize(HeaderToExcel); Finalize(TransferArr2);

          // Output Warnings
          Setlength(DataSToExcel,WarningsCount+2,2);

          for i := 1 to WarningsCount do
            DataSToExcel[i,1] := warningsArr[i];

          if warningsCount = 0 then DataSToExcel[0,1] := 'No Warnings';

          TransferArr := DataSToExcel;
          Sheet.Range[Sheet.Cells[10,1],Sheet.Cells[10+WarningsCount,2]] := TransferArr;
          Finalize(DataSToExcel);
          Finalize(TransferArr);
        end;

        //======================
        // Write "Labels" sheet
        //======================
        if LowerCase(Sheet.Name) = 'labels' then
        begin
          SetLength(HeaderToExcel,15+TotalElements+TotalNodes,52);

          // Create header data
          HeaderToExcel[0,1] := 'Created by: ' + PROG_NAME + ' by 2H Offshore Engineering';
          HeaderToExcel[1,1] := 'Version: ' + Version;
          HeaderToExcel[2,1] := 'Title: ' + Title;
          HeaderToExcel[3,1] := 'Filename: ' + filename;
          HeaderToExcel[3,5] := 'INF File: ' + infFile;
          HeaderToExcel[4,0] := 'Overview of Labels';
          HeaderToExcel[5,0] := 'Element Labels';
          HeaderToExcel[6,0] := 'Element Number';
          HeaderToExcel[6,1] := 'Line Set';
          HeaderToExcel[6,2] := 'Section Set';
          HeaderToExcel[6,3] := 'Line First Element';
          HeaderToExcel[6,4] := 'Line Last Element';
          HeaderToExcel[6,5] := 'Section First Element';
          HeaderToExcel[6,6] := 'Section Last Element';
          HeaderToExcel[6,7] := 'Node Label Adjacent Elements';

          if (TotalLabelE = 0) then
            HeaderToExcel[6,1] := 'NO ELEMENT LABEL DATA SPECIFIED (thus no entries)';

          HeaderToExcel[7+TotalElements,0] := 'Nodal Labels';
          HeaderToExcel[8+TotalElements,0] := 'Node Number';
          HeaderToExcel[8+TotalElements,1] := 'Labels';

          if (TotalLabelN = 0) then
            HeaderToExcel[8+TotalElements,1] := 'NO NODAL LABEL DATA SPECIFIED (thus no entries)';

          // Array index of start of line/section element ranges
          elemSetsIdx := TotalSingleElemLabels + 1;

          //-----------------------
          // Write Line Set column
          //-----------------------
          i := 1;
          repeat
            // Retrieve element number
            elem := eleminfo[i,1];
            elemLast := 0;
            lineIdx := 0;

            // Determine element range - only interested in searching the line/section element ranges
            for j := elemSetsIdx to TotalLabelE do
            begin
              // If start element in label index array matches element number
              if LabelElemI[j,1] = elem then
                // If last element is greater than previous last element
                if LabelElemI[j,2] > elemLast then
                begin
                  // Update last element and index
                  elemLast := LabelElemI[j,2];
                  lineIdx := j;
                end;
            end;

            // Line label found
            if lineIdx > 0 then
            begin
              // Retrieve first element number and label
              elemFirst := LabelElemI[lineIdx,1];
              elemLabel := LabelElem[lineIdx];

              // Assign line element label to element range in Excel array
              for elem := elemFirst to elemLast do
              begin
                HeaderToExcel[7+i-1,0] := elem;
                HeaderToExcel[7+i-1,1] := '{' + elemLabel + '}';
                Inc(i);
              end;
            end
            else
              Inc(i);
          until i > totalElements;

          //--------------------------
          // Write Section Set column
          //--------------------------
          i := 1;
          repeat
            // Retrieve element number
            elem := eleminfo[i,1];
            elemLast := 0;
            lineIdx := 0;

            // Retrieve element Line Set label from Excel array
            lineLabel := HeaderToExcel[7+i-1,1];

            // Strip { and } for comparison purposes
            lineLabel := Copy(lineLabel, 2, Length(lineLabel) - 2);

            // Determine element range - only interested in searching the line/section element ranges
            for j := elemSetsIdx to TotalLabelE do
            begin
              // If start element in label index array matches element number and isn't a line label
              if (LabelElemI[j,1] = elem) and (LabelElem[j] <> lineLabel) then
                // If last element is greater than previous last element
                if LabelElemI[j,2] > elemLast then
                begin
                  // Update last element and index
                  elemLast := LabelElemI[j,2];
                  lineIdx := j;
                end;
            end;

            // Section label found
            if lineIdx > 0 then
            begin
              // Retrieve first element number and label
              elemFirst := LabelElemI[lineIdx,1];
              elemLabel := LabelElem[lineIdx];

              // Assign section element label to element range in Excel array
              for elem := elemFirst to elemLast do
              begin
                HeaderToExcel[7+i-1,2] := '{' + elemLabel + '}';
                Inc(i);
              end;
            end
            else
              Inc(i);
          until i > totalElements;

          //-----------------------------------------
          // Write individual element labels columns
          //-----------------------------------------
          for i := 1 to totalElements do
          begin
            // Retrieve element number
            elem := eleminfo[i,1];
            elemLast := 0;

            // Retrieve element line and section (if exists) labels from Excel array
            lineLabel := HeaderToExcel[7+i-1,1];
            sectLabel := HeaderToExcel[7+i-1,2];

            // Strip { and } for comparison purposes
            lineLabel := Copy(lineLabel, 2, Length(lineLabel) - 2);
            sectLabel := Copy(sectLabel, 2, Length(sectLabel) - 2);

            // Initialise _Before/_After element label count
            k := 0;

            // Determine element range - only interested in searching the first/last/before/after element labels
            for j := 1 to elemSetsIdx - 1 do
            begin
              elemLabel := LabelElem[j];

              // If start element in label index array matches element number
              if LabelElemI[j,1] = elem then
              begin
                // Determine type of element: first, last, before, after
                // _First label
                if AnsiRightStr(LabelElem[j], 6) = '_First' then
                begin
                  // Strip _First
                  labelStem := AnsiLeftStr(elemLabel, AnsiPos('_First', elemLabel)-1);

                  // Assign _First label associated to a line
                  if labelStem = lineLabel then
                    HeaderToExcel[7+i-1,3] := '{' + elemLabel + '}'
                  // Assign _First label associated to a section
                  else
                    HeaderToExcel[7+i-1,5] := '{' + elemLabel + '}';
                end
                // _Last label
                else if AnsiRightStr(LabelElem[j], 5) = '_Last' then
                begin
                  // Strip _Last
                  labelStem := AnsiLeftStr(elemLabel, AnsiPos('_Last', elemLabel)-1);

                  // Assign _Last label associated to a line
                  if labelStem = lineLabel then
                    HeaderToExcel[7+i-1,4] := '{' + elemLabel + '}'
                  // Assign _Last label associated to a section
                  else
                    HeaderToExcel[7+i-1,6] := '{' + elemLabel + '}';
                end
                // _Before label (and manually assign associated _After label)
                else if AnsiRightStr(LabelElem[j], 7) = '_Before' then
                begin
                  // Strip _Before
                  labelStem := AnsiLeftStr(elemLabel, AnsiPos('_Before', elemLabel)-1);

                  // Check array index/cell not already assigned (with a _After label)
                  if HeaderToExcel[7+i-1,k+7] <> '' then Inc(k);

                  // Max columns/labels per element check
                  if  k > 99 then
                    HeaderToExcel[7+i-1,100] := 'MAX LABELS PER ELEMENT REACHED'
                  else
                  begin
                    // Assign _Before label
                    HeaderToExcel[7+i-1,k+7] := '{' + elemLabel + '}';

                    // Assign associated _After label in row below
                    HeaderToExcel[7+i,k+7] := '{' + labelStem + '_After}';
                  end;

                  Inc(k); // Next _Before/_After column
                end;
              end;
            end;
          end;

          //-----------------------
          // Write all node labels
          //-----------------------
          for i := 1 to totalNodes do
          begin
            k := 0;
            HeaderToExcel[9+totalElements+i-1,k] := IntToStr(NodeinfoI[i]);
            Inc(k);

            for j := 1 to TotalLabelN do
            begin
              if (NodeinfoI[i] >= LabelNodeI[j,1]) and (NodeInfoI[i] <= LabelNodeI[j,2]) then
              begin
                HeaderToExcel[9+TotalElements+i-1,k] := '{'+LabelNode[j]+'}'; inc(k);
              end;

              if k > 49 then
              begin
                k := 49;
                HeaderToExcel[9+TotalElements+i-1,50] := 'MAX LABELS PER NODE REACHED';
              end;
            end;
          end;

          TransferArr2 := HeaderToExcel;
          Sheet.Range[Sheet.Cells[1,1],Sheet.Cells[9+TotalElements+TotalNodes,50]] := TransferArr2;
          Finalize(HeaderToExcel);
          Finalize(TransferArr2);

          // Format sheet - autofit columns
          lastRow := Sheet.UsedRange.Rows.Count;
          lastCol := Sheet.UsedRange.Columns.Count;
          Sheet.Columns[1].EntireColumn.AutoFit;
          Sheet.Range[Sheet.Cells[7,2], Sheet.Cells[lastRow,lastCol]].Columns.AutoFit;
        end; // End of Labels sheet

        //==================================
        // Write "API-STD-2RD_Tables" sheet
        //==================================
        if LowerCase(Sheet.Name) = 'api-std-2rd_tables' then
        begin
          if APISTD2RDCodeCheck = True then  // provide the data
          begin
            MaxlinesTables := 15 + Num_Temperature*20 + 4;
            SetLength(HeaderToExcel,MaxlinesTables+1,52);

            HeaderToExcel[0,1] := 'Created by : ' + PROG_NAME + '       by 2H Offshore Engineering';
            HeaderToExcel[1,1] := 'Version    : ' + Version;
            HeaderToExcel[2,1] := 'Title      : ' + Title;
            HeaderToExcel[3,1] := 'Filename   : ' + filename;
            HeaderToExcel[3,5] := '.INF File  : ' + infFile;
            HeaderToExcel[4,5] := 'Data Echo of API-STD-2RD Tables';
            k := 5;

            if Num_Temperature > 0 then
            begin
              inc(k);
              HeaderToExcel[k,1] := ' Thermal Tables Data Echo';
              inc(k);

              for j := 1 to Num_Temperature do
              begin
                HeaderToExcel[k,0] := ' Thermal Table Number ';
                HeaderToExcel[k,1] := j;
                inc(k);
                HeaderToExcel[k,0] := 'Temperature Degree C';
                HeaderToExcel[k,1] := 'Derating Factor ';
                inc(k);

                for j2 := 1 to Round(YieldStrTempCor[j,0,1]) do
                begin
                  HeaderToExcel[k,0] := YieldStrTempCor[j,j2,1];
                  HeaderToExcel[k,1] := YieldStrTempCor[j,j2,2];
                  inc(k);
                end;
              end;
            end;

            TransferArr2 := HeaderToExcel;
            Sheet.Range[Sheet.Cells[1,1],Sheet.Cells[MaxLinesTables,50]] := TransferArr2;
            Finalize(HeaderToExcel);
            Finalize(TransferArr2);
          end;
        end; // end of API STD 2RD tables echo
      end; // end of loop over sheets

      // End of writing to workbook - update the status flag in spreadsheet
      // Delete API sheet if not used
      if APISTD2RDCodeCheck = False then
      begin
        numWorkSheets := Excel.Worksheets.Count;

        for ish := 1 to numWorkSheets do
        begin
          Sheet := Excel.Worksheets[ish];
          if (LowerCase(Sheet.Name) = 'api-std-2rd_tables') then APISTD2RDno := ish;
        end;

        Excel.DisplayAlerts := False;
        Excel.Application.WorkSheets[APISTD2RDno].Delete;
      end;

      // Refresh number of sheets
      numWorkSheets := Excel.Worksheets.Count;

      // Flag stress check is done on Cover sheet
      for i := 1 to numWorkSheets do
      begin
        Sheet := Excel.WorkSheets[i];
        if LowerCase(Sheet.Name) = 'cover' then Sheet.Cells[4,4].Value := 'Done';
      end;

      // Save and close workbook
      Excel.DisplayAlerts := False;
      Excel.Application.Workbooks[1].Save;
      WriteLn('Successfully saved Excel File - Stage 3');
      Excel.Workbooks.Close;
      Excel.IgnoreRemoteRequests := False;
      Excel.Quit;
    end;
  except on E:Exception do
  begin
    Writeln('Program has encountered error in Excel File Writing');
    Writeln('Error Message ' +E.Message+'.');
    Excel.DisplayAlerts := False;
    Excel.Workbooks.Close;
    Excel.IgnoreRemoteRequests := False;
    Excel.Quit;
    Halt(1);
  end;
  end;
end;


Procedure OutputExcel(phase : integer);
begin
try CoInitialize(nil);
  if phase = 1 then OutputExcel1;
  if phase = 2 then OutputExcel2;
  if phase = 3 then OutputExcel3;
finally
  // Required for a console application program when using OLE
  CoUninitialize;
end;
end;


{ **************************************************************************** }
//                          M A I N   P R O G R A M
{ **************************************************************************** }
begin
try
  // File handles:
  // h = .inf
  // o = .out
  // b = .tb1

  Program_Filename := paramstr(0);
  CurrentDirProgram := ExtractfilePath(Program_Filename);
  // Initialise error/warning flags
  PrintTB2 := false;
  PrintTB1 := false;
  BSIShearInput := false;
  INFnameInOutput := false;
  APISTDmethod4 := false;
  debug := False;
  IntFluidIndex := False;
  IntFluidElementOverride := False;
  pWarn := False; 
  DNVPinput := False;
  DNV_Code_Mod := True;
  ExtFluidIndex := False;
  OSTT := False;
  ExtremeData := False;
  warnFlag1 := True;
  warnTemp := True;
  WARN_ISO_TE_NEG_FLAG := false;
  WARN_DNV_TE_NEG_FLAG := false;
  WARN_BSI_STTFLAG := false;
  WARN_PIPE_SLENDERNESSFLAG := false;
  Warn_Pipe_PressureFlag := false;
  globalFactor := false;
  Functional := false;
  OutputToExcel := true; 

  // Write out header information
  WriteLn(' Program: ', PROG_NAME, ', Version: ', VERSION);

  // Tell user command line parameters and stops
  if paramcount < 2 then Help;

  // Assign filenames to variables and text files
  filename := ParamStr(1);
  infFile := ParamStr(2);
  ParamcountRemain := paramcount;

  // check the last line parameter for flags

  lastparam := ParamStr(paramcount);
  if (lastparam = '/i') or (lastparam = '/I') then
  begin;
    writeln('INF filename to be included in the output filenames');
    INFnameInOutput := true;
    ParamcountRemain := paramcount - 1;
  end;

  // Check that the .inf file exists
  if FileSearch(infFile + '.inf','') = '' then
  begin
    WriteLn(' Error: File ', infFile, '.inf not found.');
    Halt(1);
  end;

  // Code check booleans
  ISOCodeCheck:= False;
  APIRP2RDCodeCheck:= False;
  APISTD2RDCodeCheck:= False;
  DNVCodeCheck:= False;
  BSICodeCheck:= True;
  maxFileNo := 1;

  // Open the .inf file
  AssignFile(h, infFile + '.inf');
  Reset(h);
  ReadLn(h, dumStr);
  FirstWord(dumStr, tempDumpString);

  if tempDumpString = 'API' then
    begin
      WriteLn(' Error: "API" is vague. Use "API-RP-2RD" or "API-STD-2RD" instead.');
      Halt(1);
    end
  else if tempDumpString = 'API-RP-2RD' then
    begin
      APIRP2RDCodeCheck := True;
      BSICodeCheck := False;
      WriteLn(' Using API-RP-2RD Code');
      ReadLn(h, SuppElev);
    end

  else if tempDumpString = 'API-STD-2RD' then
    begin
      APISTD2RDCodeCheck := True;
      DNVCodeCheck := True;
      BSICodeCheck := False;
      WriteLn(' Using API-STD-2RD Code');

      if paramCountRemain >= 4 then
      begin
          maxFileNo := 3; filenameF := ParamStr(3);
          filenameF2 := ParamStr(4);
          if FileExists(filenameF + '.tab') = False then begin WriteLn('Error ', filenameF, '.tab not found'); halt(1); end;
          if FileExists(filenameF2 + '.tab') = False then begin WriteLn('Error ', filenameF2, '.tab not found'); halt(1); end;
      end;

      if paramcountRemain = 3 then
      begin
          if UpperCase(ParamStr(3)) = 'FUNCTIONAL' then
          begin
              maxFileNo := 1;
              Functional := true;
          end
          else
          begin
              maxFileNo := 2; filenameF := ParamStr(3);
              if FileExists(filenameF + '.tab') = False then
              begin
                  WriteLn('Error ', filenameF, '.tab not found');
                  halt(1);
              end;
          end;
      end;

      if (maxFileNo = 1) and (Functional = false) then
      begin
          WriteLn(' User is required for a DNV code check which is functional check to include Functional as 3rd command line parameter');
          Halt(1);
      end;
      ReadLn(h, SuppElev);
    end
  // 1.51: ISO
  else if tempDumpString = 'ISO' then
    begin
      ISOCodeCheck := True;
      BSICodeCheck := False;
      WriteLn(' Using ISO Code');
      ReadLn(h, SuppElev);
    end

  // ISO: 1.51
  else if (tempDumpString = 'DNV') or (tempDumpstring = 'DNV-OS-F201') or (tempDumpstring = 'DNV-OS-F201-TL') then
    begin
      DNVCodeCheck := True;
      BSICodeCheck := False;
      if tempDumpString = 'DNV-OS-F201-TL' then DNV_Code_Mod := false;
      WriteLn( 'Using DNV-OS-F201 Code Check');
      APIRP2RDcodeCheck := True;
      WriteLn( 'Secondary Code Check is API-RP-2RD');

      if paramCountRemain >= 4 then
        begin
          maxFileNo := 3; filenameF := ParamStr(3);
          filenameF2 := ParamStr(4);
          if FileExists(filenameF + '.tab') = False then begin WriteLn('Error ', filenameF, '.tab not found'); halt(1); end;
          if FileExists(filenameF2 + '.tab') = False then begin WriteLn('Error ', filenameF2, '.tab not found'); halt(1); end;
        end;

      if paramcountRemain = 3 then begin
        if UpperCase(ParamStr(3)) = 'FUNCTIONAL' then
        begin
            maxFileNo := 1;
            Functional := true;
        end
        else
        begin
            maxFileNo := 2; filenameF := ParamStr(3);
            if FileExists(filenameF + '.tab') = False then
            begin
                WriteLn('Error ', filenameF, '.tab not found');
                halt(1);
            end;
        end;
      end;

      if (maxFileNo = 1) and (Functional = false) then
      begin
          WriteLn(' User is required for a DNV code check which is functional check to include Functional as 3rd command line parameter');
          Halt(1);
      end;
      ReadLn(h, SuppElev);
    end

  else
    begin
      WriteLn(' Using BS Code');
      Val(TempDumpString, SuppElev, Code);
      if (Code <> 0) then ReadLn(h, SuppElev);
    end;

  // Read in the rest of the .inf file header data
  ReadLn(h, tubZeroTens);
  ReadLn(h, intCoat);
  CloseFile(h);

  // Check whether coating Od is in INF file
  dumstr := FindLine(h, 'COATING OD');
  if dumStr = 'COATING OD' then CoatODFlag := true;
  // Check whether .tb1 file is requested.
  dumstr := FindLine(h, 'PRINT TB1');
  if dumstr = 'PRINT TB1' then begin;
   Writeln(' Printing of TB1 requested');
   PRINTTB1 := true;
  end;
  // Check whether .tb2 file is to be created
  dumStr := FindLine(h, 'PRINT TB2');
  if dumStr = 'PRINT TB2' then begin;
    WriteLn(' Printing of TB2 requested');
    PRINTTB2 := True;
  end;
  dumstr := FindLine(h, 'BSISHEARINPUT');
  // Check whether BSI Shearinput
  if dumstr = 'BSISHEARINPUT' then begin;
    Writeln(' BSI Shear and Torque input expected in STT loads');
    BSIShearInput := true;
  end;
  dumstr := FindLine(h, 'APISTD2RD METHOD4');
  if dumstr = 'APISTD2RD METHOD4' then begin;
    Writeln(' API STD 2RD Method 4 - Activation calculation of bending strain from curvature ');
    APISTDmethod4 := true;
    end;
  // Check whether to output to excel
  dumstr := Findline(h,'NO EXCEL');
  if dumstr = 'NO EXCEL' then begin;
   Writeln(' No excel file output will be produced');
   OutputToExcel := false;
   end;
  dumstr := Findline(h,'EXTREME RESPONSE');
  if dumstr = 'EXTREME RESPONSE' then
   begin;
    Writeln(' Extreme response rather envelope data is expected');
    ExtremeData := True;
   end;

  // Allow users to switch on debug options in inf file
  dumStr := FindLine(h, 'TURN ON DEBUG');

  if dumStr = 'TURN ON DEBUG' then debug := True;
  if debug = True then WriteLn(' *** Debug is turned ON ***' );

  // Is PRES DNV a keyword
  dumStr := FindLine(h, 'PRES DNV');

  if (dumStr <> NOT_FOUND) then DNVPinput := True;

  // Check that the .out file exists
  if FileSearch(filename + '.out','')='' then begin
    WriteLn (' Error: File ', filename, '.out not found.');
    Halt(1);
  end;

  // Check that the .tab file exists
  // Note version 8 and greater .tab files may have a -PST suffix
  if ((FileSearch(filename + '.tab','') = '')
    and (FileSearch(filename + '-pst.tab','') = '')) then begin
    WriteLn (' Error: File ', filename, '.tab not found.');
    Halt(1);
  end;

  // Set up .out file to read from
  AssignFile(o, filename + '.out');

  // Set up .tab file to read from (.tab file may contain a -pst suffix)
  AssignTabFile(filename);

  // Convert filename to uppercase
  filename := UpperCase(filename);

  // Determine Flexcom version
  GetFlexcomVersion(o, flexVer1, flexVer2);

  // Read title from .out file
  title := FindLine(o, 'ANALYSIS TITLE');
  for i := 1 to 3 do ReadLn(o, title);
  title := Trim(title);

  WriteLn('                             Reading Element Data');
  // Read number nodes, elements, linear flex joints
  GetModelDetails;

  // Write element and node number in left margin of pages
  WriteElem;

  WriteLn('                             Writing Displacement Table');
  // Write displacement table
  Displacements;

  WriteLn('                             Reading Geometry');
  // Read element connectivity and properties
  Geometry;
  Writeln('                             Reading Label Data');
  ReadLabels;
  if APISTD2RDCodeCheck = true then CheckValidMethod4;

  WriteLn('                             Reading Factors');
  FactorsInput;

  WriteLn('                             Calculating Tension Correction');
  // Calculate and write tension correction
  TensionCorrection;

  WriteLn ('                             Reading Forces');
  // Write forces
  Forces;

  if STT and (FileSearch(filename + '.grd','') <> '') then
    begin
      WriteLn('                             Extracting Time Domain Max Stress Loads');
      // Determines time domain loads to be used in stress calc
      STTLoads;
      if OSTT = true then
      begin;
      Writeln('                             Compiling Time Domain Stress Output');
      CompileTb3;
      end;
    end
  else if STT then
    begin
      WriteLn('                     *** GRD not present ***');
      halt(1);
    end;

  if DNVCodeCheck then begin
    WriteLn('                             DNV Code Checking');
    DNVCal;
  end;

  if APIRP2RDCodeCheck then begin
    WriteLn('                             API-RP-2RD Checking');
  end;

  WriteLn('                             Calculating VM Stresses');
  // Calculates von Mises stresses
  if not STT then CalStress;

  // Compile .tb2 file if requested
  if PRINTTB2 then CompileTB2;
  WriteLn('                             Compiling Springs and Flex Joints Table ');

  // Calculate spring forces
  Springs;
  WriteLn('                             Springs Complete');

  // Calculate articulations / flex joints angles
  FlexJoints;
  WriteLn('                             Flex Joints Complete');
  // Close .tab file
  CloseFile(t);
  // Close .out file
  CloseFile(o);

  // Compile .tb1 file by merging forces, displacement and factors
  if PRINTTB1 = True then
  begin;
    WriteLn('                             Compiling TB1 File');
    CompileTB1;
    WriteLn('                             TB1 File Complete');

    if APISTD2RDCodeCheck = True then
    begin;
      CompileTB4TB5;
      Writeln('                             TB4 File Complete');
    end;
  end;

  if OutputToExcel = true then OutputExcel(1);
  if OutputToExcel = true then OutputExcel(2);
  if OutputToExcel = true then OutputExcel(3);

//---------------
// Error handler
//---------------
except on E: Exception do
begin
  WriteLn('An error occured: ' + E.Message + '.');
end;
end;
end.
