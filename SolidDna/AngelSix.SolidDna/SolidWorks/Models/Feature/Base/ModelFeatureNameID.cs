using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents the suppression state action of a <see cref="ModelFeature"/>
    /// 
    /// NOTE: Known types are here http://help.solidworks.com/2019/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swFeatureSuppressionAction_e.html
    /// 
    /// </summary>
    public enum ModelFeatureNameID
    {
        /// <summary>
        /// Absolute view
        /// </summary>
        AbsoluteView = swFeatureNameID_e.swFmAbsoluteView,

        /// <summary>
        /// 3D contact
        /// </summary>
        Contact3D = swFeatureNameID_e.swFmAEM3DContact,

        /// <summary>
        /// Gravity
        /// </summary>
        Gravity = swFeatureNameID_e.swFmAEMGravity,

        /// <summary>
        /// Linear damper
        /// </summary>
        LinearDamper = swFeatureNameID_e.swFmAEMLinearDamper,

        /// <summary>
        /// Linear force
        /// </summary>
        LinearForce = swFeatureNameID_e.swFmAEMLinearForce,

        /// <summary>
        /// Linear motion spring
        /// </summary>
        LinearMotionSpring = swFeatureNameID_e.swFmAEMLinearMotionSpring,

        /// <summary>
        /// Linear motor
        /// </summary>
        LinearMotor = swFeatureNameID_e.swFmAEMLinearMotor,

        /// <summary>
        /// Linear spring
        /// </summary>
        LinearSpring = swFeatureNameID_e.swFmAEMLinearSpring,

        /// <summary>
        /// Rotational motor
        /// </summary>
        RotationalMotor = swFeatureNameID_e.swFmAEMRotationalMotor,

        /// <summary>
        /// Torque
        /// </summary>
        Torque = swFeatureNameID_e.swFmAEMTorque,

        /// <summary>
        /// Torsional damper
        /// </summary>
        TorsionalDamper = swFeatureNameID_e.swFmAEMTorsionalDamper,

        /// <summary>
        /// Torsional motion spring
        /// </summary>
        TorsionalMotionSpring = swFeatureNameID_e.swFmAEMTorsionalMotionSpring,

        /// <summary>
        /// Torsional spring
        /// </summary>
        TorsionalSpring = swFeatureNameID_e.swFmAEMTorsionalSpring,

        /// <summary>
        /// Attribute
        /// </summary>
        Attribute = swFeatureNameID_e.swFmAttribute,

        /// <summary>
        /// Auxiliary view
        /// </summary>
        AuxiliaryView = swFeatureNameID_e.swFmAuxiliaryView,

        /// <summary>
        /// Base body
        /// </summary>
        BaseBody = swFeatureNameID_e.swFmBaseBody,

        /// <summary>
        /// Base flange
        /// </summary>
        BaseFlange = swFeatureNameID_e.swFmBaseFlange,

        /// <summary>
        /// Belt and chain
        /// </summary>
        BeltAndChain = swFeatureNameID_e.swFmBeltAndChain,

        /// <summary>
        /// Blend
        /// </summary>
        Blend = swFeatureNameID_e.swFmBlend,

        /// <summary>
        /// Blend cut
        /// </summary>
        BlendCut = swFeatureNameID_e.swFmBlendCut,

        /// <summary>
        /// BOM table feature
        /// </summary>
        BomTableFeature = swFeatureNameID_e.swFmBomTableFeature,

        /// <summary>
        /// Boss
        /// </summary>
        Boss = swFeatureNameID_e.swFmBoss,

        /// <summary>
        /// BoundingBox
        /// </summary>
        BoundingBox = swFeatureNameID_e.swFmBoundingBox,

        /// <summary>
        /// Break corner
        /// </summary>
        BreakCorner = swFeatureNameID_e.swFmBreakCorner,

        /// <summary>
        /// Cavity
        /// </summary>
        Cavity = swFeatureNameID_e.swFmCavity,

        /// <summary>
        /// Center mark
        /// </summary>
        CenterMark = swFeatureNameID_e.swFmCenterMark,

        /// <summary>
        /// Chamfer
        /// </summary>
        Chamfer = swFeatureNameID_e.swFmChamfer,

        /// <summary>
        /// Circular pattern
        /// </summary>
        CircularPattern = swFeatureNameID_e.swFmCirPattern,

        /// <summary>
        /// Coordinate system
        /// </summary>
        CoordinateSystem = swFeatureNameID_e.swFmCoordinateSystem,

        /// <summary>
        /// Sheet metal corner relief
        /// </summary>
        SheetMetalCornerRelief = swFeatureNameID_e.swFmCornerRelief,

        /// <summary>
        /// Corner trim
        /// </summary>
        CornerTrim = swFeatureNameID_e.swFmCornerTrim,

        /// <summary>
        /// Cosmetic thread
        /// </summary>
        CosmeticThread = swFeatureNameID_e.swFmCosmeticThread,

        /// <summary>
        /// Cross break
        /// </summary>
        CrossBreak = swFeatureNameID_e.swFmCrossBreak,

        /// <summary>
        /// CurvePattern
        /// </summary>
        CurvePattern = swFeatureNameID_e.swFmCurvePattern,

        /// <summary>
        /// Cut
        /// </summary>
        Cut = swFeatureNameID_e.swFmCut,

        /// <summary>
        /// Derived pattern
        /// </summary>
        DerivedPattern = swFeatureNameID_e.swFmDerivedLPattern,

        /// <summary>
        /// Detail circle
        /// </summary>
        DetailCircle = swFeatureNameID_e.swFmDetailCircle,

        /// <summary>
        /// Detail view
        /// </summary>
        DetailView = swFeatureNameID_e.swFmDetailView,

        /// <summary>
        /// DimPattern
        /// </summary>
        DimPattern = swFeatureNameID_e.swFmDimPattern,

        /// <summary>
        /// Draft
        /// </summary>
        Draft = swFeatureNameID_e.swFmDraft,

        /// <summary>
        /// Drawing section line
        /// </summary>
        DrawingSectionLine = swFeatureNameID_e.swFmDrSectionLine,

        /// <summary>
        /// Drawing sheet
        /// </summary>
        DrawingSheet = swFeatureNameID_e.swFmDrSheet,

        /// <summary>
        /// Edge flange
        /// </summary>
        EdgeFlange = swFeatureNameID_e.swFmEdgeFlange,

        /// <summary>
        /// Extrusion
        /// </summary>
        Extrusion = swFeatureNameID_e.swFmExtrusion,

        /// <summary>
        /// Feature folder
        /// </summary>
        FeatureFolder = swFeatureNameID_e.swFmFeatureFolder,

        /// <summary>
        /// Fillet
        /// </summary>
        Fillet = swFeatureNameID_e.swFmFillet,

        /// <summary>
        /// APattern
        /// </summary>
        FillPattern = swFeatureNameID_e.swFmFillPattern,

        /// <summary>
        /// Flat pattern
        /// </summary>
        FlatPattern = swFeatureNameID_e.swFmFlatPattern,

        /// <summary>
        /// Flatten bends
        /// </summary>
        FlattenBends = swFeatureNameID_e.swFmFlattenBends,

        /// <summary>
        /// Form tool instance
        /// </summary>
        FormToolInstance = swFeatureNameID_e.swFmFormToolInstance,

        /// <summary>
        /// GroundPlane
        /// </summary>
        GroundPlane = swFeatureNameID_e.swFmGroundPlane,

        /// <summary>
        /// Hole table feature
        /// </summary>
        HoleTableFeature = swFeatureNameID_e.swFmHoleTableFeature,

        /// <summary>
        /// Hole wizard
        /// </summary>
        HoleWizard = swFeatureNameID_e.swFmHoleWzd,

        /// <summary>
        /// Imported
        /// </summary>
        Imported = swFeatureNameID_e.swFmImported,

        /// <summary>
        /// Library feature
        /// </summary>
        LibraryFeature = swFeatureNameID_e.swFmLibraryFeature,

        /// <summary>
        /// LocalChainPattern
        /// </summary>
        LocalChainPattern = swFeatureNameID_e.swFmLocalChainPattern,

        /// <summary>
        /// LocalCirPattern
        /// </summary>
        LocalCircularPattern = swFeatureNameID_e.swFmLocalCirPattern,

        /// <summary>
        /// LocalCurvePattern
        /// </summary>
        LocalCurvePattern = swFeatureNameID_e.swFmLocalCurvePattern,

        /// <summary>
        /// LocalLPattern
        /// </summary>
        LocalLinearPattern = swFeatureNameID_e.swFmLocalLPattern,

        /// <summary>
        /// LocalSketchPattern
        /// </summary>
        LocalSketchPattern = swFeatureNameID_e.swFmLocalSketchPattern,

        /// <summary>
        /// Linear pattern
        /// </summary>
        LinearPattern = swFeatureNameID_e.swFmLPattern,

        /// <summary>
        /// Tangent mate
        /// </summary>
        MateCamTangent = swFeatureNameID_e.swFmMateCamTangent,

        /// <summary>
        /// Coincident mate
        /// </summary>
        MateCoincident = swFeatureNameID_e.swFmMateCoincident,

        /// <summary>
        /// Concentric mate
        /// </summary>
        MateConcentric = swFeatureNameID_e.swFmMateConcentric,

        /// <summary>
        /// Mate controller
        /// </summary>
        // TODO: Add after upgrading to SOLIDWORKS 2023 binary reference files
        //MateController = swFeatureNameID_e.swFmMateController,

        /// <summary>
        /// Distance mate
        /// </summary>
        MateDistanceDim = swFeatureNameID_e.swFmMateDistanceDim,

        /// <summary>
        /// Gear mate
        /// </summary>
        MateGearDim = swFeatureNameID_e.swFmMateGearDim,

        /// <summary>
        /// Hinge mate
        /// </summary>
        MateHinge = swFeatureNameID_e.swFmMateHinge,

        /// <summary>
        /// Linear coupler mate
        /// </summary>
        MateLinearCoupler = swFeatureNameID_e.swFmMateLinearCoupler,

        /// <summary>
        /// Lock mate
        /// </summary>
        MateLock = swFeatureNameID_e.swFmMateLock,

        /// <summary>
        /// Parallel mate
        /// </summary>
        MateParallel = swFeatureNameID_e.swFmMateParallel,

        /// <summary>
        /// Perpendicular mate
        /// </summary>
        MatePerpendicular = swFeatureNameID_e.swFmMatePerpendicular,

        /// <summary>
        /// Angle mate
        /// </summary>
        MatePlanarAngle = swFeatureNameID_e.swFmMatePlanarAngleDim,

        /// <summary>
        /// Profile center mate
        /// </summary>
        MateProfileCenter = swFeatureNameID_e.swFmMateProfileCenter,

        /// <summary>
        /// Rack and pinion mate
        /// </summary>
        MateRackAndPinion = swFeatureNameID_e.swFmMateRackPinionDim,

        /// <summary>
        /// Screw mate
        /// </summary>
        MateScrew = swFeatureNameID_e.swFmMateScrew,

        /// <summary>
        /// Slot mate
        /// </summary>
        MateSlot = swFeatureNameID_e.swFmMateSlot,

        /// <summary>
        /// Symmetric mate
        /// </summary>
        MateSymmetric = swFeatureNameID_e.swFmMateSymmetric,

        /// <summary>
        /// Tangent mate
        /// </summary>
        MateTangent = swFeatureNameID_e.swFmMateTangent,

        /// <summary>
        /// Universal joint mate
        /// </summary>
        MateUniversalJoint = swFeatureNameID_e.swFmMateUniversalJoint,

        /// <summary>
        /// Width mate
        /// </summary>
        MateWidth = swFeatureNameID_e.swFmMateWidth,

        /// <summary>
        /// Mirror component
        /// </summary>
        MirrorComponent = swFeatureNameID_e.swFmMirrorComponent,

        /// <summary>
        /// Mirror pattern
        /// </summary>
        MirrorPattern = swFeatureNameID_e.swFmMirrorPattern,

        /// <summary>
        /// Mirror solid
        /// </summary>
        MirrorSolid = swFeatureNameID_e.swFmMirrorSolid,

        /// <summary>
        /// Sheet metal normal cut
        /// </summary>
        SheetMetalNormalCut = swFeatureNameID_e.swFmNormalCut,

        /// <summary>
        /// One bend
        /// </summary>
        OneBend = swFeatureNameID_e.swFmOneBend,

        /// <summary>
        /// Process bends
        /// </summary>
        ProcessBends = swFeatureNameID_e.swFmProcessBends,

        /// <summary>
        /// Profile feature
        /// </summary>
        ProfileFeature = swFeatureNameID_e.swFmProfileFeature,

        /// <summary>
        /// Reference axis
        /// </summary>
        ReferenceAxis = swFeatureNameID_e.swFmRefAxis,

        /// <summary>
        /// Reference curve
        /// </summary>
        RefCurve = swFeatureNameID_e.swFmRefCurve,

        /// <summary>
        /// Reference
        /// </summary>
        Reference = swFeatureNameID_e.swFmReference,

        /// <summary>
        /// Reference curve
        /// </summary>
        ReferenceCurve = swFeatureNameID_e.swFmReferenceCurve,

        /// <summary>
        /// Reference plane
        /// </summary>
        ReferencePlane = swFeatureNameID_e.swFmRefPlane,

        /// <summary>
        /// Reference or sweep surface
        /// </summary>
        ReferenceOrSweepSurface = swFeatureNameID_e.swFmRefSurface,

        /// <summary>
        /// Relative view
        /// </summary>
        RelativeView = swFeatureNameID_e.swFmRelativeView,

        /// <summary>
        /// Reference cut
        /// </summary>
        ReferenceCut = swFeatureNameID_e.swFmRevCut,

        /// <summary>
        /// Revision table feature
        /// </summary>
        RevisionTableFeature = swFeatureNameID_e.swFmRevisionTableFeature,

        /// <summary>
        /// Revolve
        /// </summary>
        Revolution = swFeatureNameID_e.swFmRevolution,

        /// <summary>
        /// Section view of assembly
        /// </summary>
        SectionAssemblyView = swFeatureNameID_e.swFmSectionAssemView,

        /// <summary>
        /// Section view of part
        /// </summary>
        SectionPartView = swFeatureNameID_e.swFmSectionPartView,

        /// <summary>
        /// Sheet metal
        /// </summary>
        SheetMetal = swFeatureNameID_e.swFmSheetMetal,

        /// <summary>
        /// Shell
        /// </summary>
        Shell = swFeatureNameID_e.swFmShell,

        /// <summary>
        /// Sketch bend
        /// </summary>
        SketchBend = swFeatureNameID_e.swFmSketchBend,

        /// <summary>
        /// Sketch hole
        /// </summary>
        SketchHole = swFeatureNameID_e.swFmSketchHole,

        /// <summary>
        /// SketchPattern
        /// </summary>
        SketchPattern = swFeatureNameID_e.swFmSketchPattern,

        /// <summary>
        /// Sheet metal 3D bend
        /// </summary>
        SheetMetal3DBend = swFeatureNameID_e.swFmSM3dBend,

        /// <summary>
        /// Sheet metal gusset
        /// </summary>
        SheetMetalGusset = swFeatureNameID_e.swFmSMGusset,

        /// <summary>
        /// Solid body folder
        /// </summary>
        SolidBodyFolder = swFeatureNameID_e.swFmSolidBodyFolder,

        /// <summary>
        /// Stock
        /// </summary>
        Stock = swFeatureNameID_e.swFmStock,

        /// <summary>
        /// Structure system corner feature
        /// </summary>
        StructureSystemCornerFeature = swFeatureNameID_e.swFmStrctSysCnrFeat,

        /// <summary>
        /// Structure system corner group feature
        /// </summary>
        StructureSystemCornerGroupFeature = swFeatureNameID_e.swFmStrctSysCnrGrpFeat,

        /// <summary>
        /// Structure system corner management feature
        /// </summary>
        StructureSystemCornerManagementFeature = swFeatureNameID_e.swFmStrctSysCnrMgmtFeat,

        /// <summary>
        /// Structure system feature
        /// </summary>
        StructureSystemFeature = swFeatureNameID_e.swFmStrctSysFeat,

        /// <summary>
        /// Structure system group feature
        /// </summary>
        StructureSystemGroupFeature = swFeatureNameID_e.swFmStrctSysGrpFeat,

        /// <summary>
        /// Structure system member feature
        /// </summary>
        StructureSystemMemberFeature = swFeatureNameID_e.swFmStrctSysMbrFeat,

        /// <summary>
        /// Surface body folder
        /// </summary>
        SurfaceBodyFolder = swFeatureNameID_e.swFmSurfaceBodyFolder,

        /// <summary>
        /// Surface cut
        /// </summary>
        SurfCut = swFeatureNameID_e.swFmSurfCut,

        /// <summary>
        /// Sweep boss
        /// </summary>
        FmSweep = swFeatureNameID_e.swFmSweep,

        /// <summary>
        /// Sweep cut
        /// </summary>
        SweepCut = swFeatureNameID_e.swFmSweepCut,

        /// <summary>
        /// Sweep thread
        /// </summary>
        SweepThread = swFeatureNameID_e.swFmSweepThread,

        /// <summary>
        /// Swept flange
        /// </summary>
        SweptFlange = swFeatureNameID_e.swFmSweptFlange,

        /// <summary>
        /// Tab and Slot
        /// </summary>
        TabAndSlot = swFeatureNameID_e.swFmTabAndSlot,

        /// <summary>
        /// TablePattern
        /// </summary>
        TablePattern = swFeatureNameID_e.swFmTablePattern,

        /// <summary>
        /// Thicken
        /// </summary>
        Thicken = swFeatureNameID_e.swFmThicken,

        /// <summary>
        /// Thicken cut
        /// </summary>
        ThickenCut = swFeatureNameID_e.swFmThickenCut,

        /// <summary>
        /// Unfolded view
        /// </summary>
        UnfoldedView = swFeatureNameID_e.swFmUnfoldedView,

        /// <summary>
        /// Variable fillet
        /// </summary>
        VariableFillet = swFeatureNameID_e.swFmVarFillet,

        /// <summary>
        /// Weldment feature
        /// </summary>
        WeldMemberFeat = swFeatureNameID_e.swFmWeldMemberFeat
    }
}
