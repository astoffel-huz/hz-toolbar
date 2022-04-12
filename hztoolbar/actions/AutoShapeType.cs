#nullable enable

using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Drawing;

namespace hztoolbar.actions
{
    public class ChangeAutoShapeType : ToolbarAction
    {

        private static ImmutableDictionary<string, Office.MsoAutoShapeType> SHAPES = new Dictionary<string, Office.MsoAutoShapeType>
        {
            ["Mixed"] = Office.MsoAutoShapeType.msoShapeMixed,
            ["Rectangle"] = Office.MsoAutoShapeType.msoShapeRectangle,
            ["Parallelogram"] = Office.MsoAutoShapeType.msoShapeParallelogram,
            ["Trapezoid"] = Office.MsoAutoShapeType.msoShapeTrapezoid,
            ["Diamond"] = Office.MsoAutoShapeType.msoShapeDiamond,
            ["RoundedRectangle"] = Office.MsoAutoShapeType.msoShapeRoundedRectangle,
            ["Octagon"] = Office.MsoAutoShapeType.msoShapeOctagon,
            ["IsoscelesTriangle"] = Office.MsoAutoShapeType.msoShapeIsoscelesTriangle,
            ["RightTriangle"] = Office.MsoAutoShapeType.msoShapeRightTriangle,
            ["Oval"] = Office.MsoAutoShapeType.msoShapeOval,
            ["Hexagon"] = Office.MsoAutoShapeType.msoShapeHexagon,
            ["Cross"] = Office.MsoAutoShapeType.msoShapeCross,
            ["RegularPentagon"] = Office.MsoAutoShapeType.msoShapeRegularPentagon,
            ["Can"] = Office.MsoAutoShapeType.msoShapeCan,
            ["Cube"] = Office.MsoAutoShapeType.msoShapeCube,
            ["Bevel"] = Office.MsoAutoShapeType.msoShapeBevel,
            ["FoldedCorner"] = Office.MsoAutoShapeType.msoShapeFoldedCorner,
            ["SmileyFace"] = Office.MsoAutoShapeType.msoShapeSmileyFace,
            ["Donut"] = Office.MsoAutoShapeType.msoShapeDonut,
            ["NoSymbol"] = Office.MsoAutoShapeType.msoShapeNoSymbol,
            ["BlockArc"] = Office.MsoAutoShapeType.msoShapeBlockArc,
            ["Heart"] = Office.MsoAutoShapeType.msoShapeHeart,
            ["LightningBolt"] = Office.MsoAutoShapeType.msoShapeLightningBolt,
            ["Sun"] = Office.MsoAutoShapeType.msoShapeSun,
            ["Moon"] = Office.MsoAutoShapeType.msoShapeMoon,
            ["Arc"] = Office.MsoAutoShapeType.msoShapeArc,
            ["DoubleBracket"] = Office.MsoAutoShapeType.msoShapeDoubleBracket,
            ["DoubleBrace"] = Office.MsoAutoShapeType.msoShapeDoubleBrace,
            ["Plaque"] = Office.MsoAutoShapeType.msoShapePlaque,
            ["LeftBracket"] = Office.MsoAutoShapeType.msoShapeLeftBracket,
            ["RightBracket"] = Office.MsoAutoShapeType.msoShapeRightBracket,
            ["LeftBrace"] = Office.MsoAutoShapeType.msoShapeLeftBrace,
            ["RightBrace"] = Office.MsoAutoShapeType.msoShapeRightBrace,
            ["RightArrow"] = Office.MsoAutoShapeType.msoShapeRightArrow,
            ["LeftArrow"] = Office.MsoAutoShapeType.msoShapeLeftArrow,
            ["UpArrow"] = Office.MsoAutoShapeType.msoShapeUpArrow,
            ["DownArrow"] = Office.MsoAutoShapeType.msoShapeDownArrow,
            ["LeftRightArrow"] = Office.MsoAutoShapeType.msoShapeLeftRightArrow,
            ["UpDownArrow"] = Office.MsoAutoShapeType.msoShapeUpDownArrow,
            ["QuadArrow"] = Office.MsoAutoShapeType.msoShapeQuadArrow,
            ["LeftRightUpArrow"] = Office.MsoAutoShapeType.msoShapeLeftRightUpArrow,
            ["BentArrow"] = Office.MsoAutoShapeType.msoShapeBentArrow,
            ["UTurnArrow"] = Office.MsoAutoShapeType.msoShapeUTurnArrow,
            ["LeftUpArrow"] = Office.MsoAutoShapeType.msoShapeLeftUpArrow,
            ["BentUpArrow"] = Office.MsoAutoShapeType.msoShapeBentUpArrow,
            ["CurvedRightArrow"] = Office.MsoAutoShapeType.msoShapeCurvedRightArrow,
            ["CurvedLeftArrow"] = Office.MsoAutoShapeType.msoShapeCurvedLeftArrow,
            ["CurvedUpArrow"] = Office.MsoAutoShapeType.msoShapeCurvedUpArrow,
            ["CurvedDownArrow"] = Office.MsoAutoShapeType.msoShapeCurvedDownArrow,
            ["StripedRightArrow"] = Office.MsoAutoShapeType.msoShapeStripedRightArrow,
            ["NotchedRightArrow"] = Office.MsoAutoShapeType.msoShapeNotchedRightArrow,
            ["Pentagon"] = Office.MsoAutoShapeType.msoShapePentagon,
            ["Chevron"] = Office.MsoAutoShapeType.msoShapeChevron,
            ["RightArrowCallout"] = Office.MsoAutoShapeType.msoShapeRightArrowCallout,
            ["LeftArrowCallout"] = Office.MsoAutoShapeType.msoShapeLeftArrowCallout,
            ["UpArrowCallout"] = Office.MsoAutoShapeType.msoShapeUpArrowCallout,
            ["DownArrowCallout"] = Office.MsoAutoShapeType.msoShapeDownArrowCallout,
            ["LeftRightArrowCallout"] = Office.MsoAutoShapeType.msoShapeLeftRightArrowCallout,
            ["UpDownArrowCallout"] = Office.MsoAutoShapeType.msoShapeUpDownArrowCallout,
            ["QuadArrowCallout"] = Office.MsoAutoShapeType.msoShapeQuadArrowCallout,
            ["CircularArrow"] = Office.MsoAutoShapeType.msoShapeCircularArrow,
            ["FlowchartProcess"] = Office.MsoAutoShapeType.msoShapeFlowchartProcess,
            ["FlowchartAlternateProcess"] = Office.MsoAutoShapeType.msoShapeFlowchartAlternateProcess,
            ["FlowchartDecision"] = Office.MsoAutoShapeType.msoShapeFlowchartDecision,
            ["FlowchartData"] = Office.MsoAutoShapeType.msoShapeFlowchartData,
            ["FlowchartPredefinedProcess"] = Office.MsoAutoShapeType.msoShapeFlowchartPredefinedProcess,
            ["FlowchartInternalStorage"] = Office.MsoAutoShapeType.msoShapeFlowchartInternalStorage,
            ["FlowchartDocument"] = Office.MsoAutoShapeType.msoShapeFlowchartDocument,
            ["FlowchartMultidocument"] = Office.MsoAutoShapeType.msoShapeFlowchartMultidocument,
            ["FlowchartTerminator"] = Office.MsoAutoShapeType.msoShapeFlowchartTerminator,
            ["FlowchartPreparation"] = Office.MsoAutoShapeType.msoShapeFlowchartPreparation,
            ["FlowchartManualInput"] = Office.MsoAutoShapeType.msoShapeFlowchartManualInput,
            ["FlowchartManualOperation"] = Office.MsoAutoShapeType.msoShapeFlowchartManualOperation,
            ["FlowchartConnector"] = Office.MsoAutoShapeType.msoShapeFlowchartConnector,
            ["FlowchartOffpageConnector"] = Office.MsoAutoShapeType.msoShapeFlowchartOffpageConnector,
            ["FlowchartCard"] = Office.MsoAutoShapeType.msoShapeFlowchartCard,
            ["FlowchartPunchedTape"] = Office.MsoAutoShapeType.msoShapeFlowchartPunchedTape,
            ["FlowchartSummingJunction"] = Office.MsoAutoShapeType.msoShapeFlowchartSummingJunction,
            ["FlowchartOr"] = Office.MsoAutoShapeType.msoShapeFlowchartOr,
            ["FlowchartCollate"] = Office.MsoAutoShapeType.msoShapeFlowchartCollate,
            ["FlowchartSort"] = Office.MsoAutoShapeType.msoShapeFlowchartSort,
            ["FlowchartExtract"] = Office.MsoAutoShapeType.msoShapeFlowchartExtract,
            ["FlowchartMerge"] = Office.MsoAutoShapeType.msoShapeFlowchartMerge,
            ["FlowchartStoredData"] = Office.MsoAutoShapeType.msoShapeFlowchartStoredData,
            ["FlowchartDelay"] = Office.MsoAutoShapeType.msoShapeFlowchartDelay,
            ["FlowchartSequentialAccessStorage"] = Office.MsoAutoShapeType.msoShapeFlowchartSequentialAccessStorage,
            ["FlowchartMagneticDisk"] = Office.MsoAutoShapeType.msoShapeFlowchartMagneticDisk,
            ["FlowchartDirectAccessStorage"] = Office.MsoAutoShapeType.msoShapeFlowchartDirectAccessStorage,
            ["FlowchartDisplay"] = Office.MsoAutoShapeType.msoShapeFlowchartDisplay,
            ["Explosion1"] = Office.MsoAutoShapeType.msoShapeExplosion1,
            ["Explosion2"] = Office.MsoAutoShapeType.msoShapeExplosion2,
            ["4pointStar"] = Office.MsoAutoShapeType.msoShape4pointStar,
            ["5pointStar"] = Office.MsoAutoShapeType.msoShape5pointStar,
            ["8pointStar"] = Office.MsoAutoShapeType.msoShape8pointStar,
            ["16pointStar"] = Office.MsoAutoShapeType.msoShape16pointStar,
            ["24pointStar"] = Office.MsoAutoShapeType.msoShape24pointStar,
            ["32pointStar"] = Office.MsoAutoShapeType.msoShape32pointStar,
            ["UpRibbon"] = Office.MsoAutoShapeType.msoShapeUpRibbon,
            ["DownRibbon"] = Office.MsoAutoShapeType.msoShapeDownRibbon,
            ["CurvedUpRibbon"] = Office.MsoAutoShapeType.msoShapeCurvedUpRibbon,
            ["CurvedDownRibbon"] = Office.MsoAutoShapeType.msoShapeCurvedDownRibbon,
            ["VerticalScroll"] = Office.MsoAutoShapeType.msoShapeVerticalScroll,
            ["HorizontalScroll"] = Office.MsoAutoShapeType.msoShapeHorizontalScroll,
            ["Wave"] = Office.MsoAutoShapeType.msoShapeWave,
            ["DoubleWave"] = Office.MsoAutoShapeType.msoShapeDoubleWave,
            ["RectangularCallout"] = Office.MsoAutoShapeType.msoShapeRectangularCallout,
            ["RoundedRectangularCallout"] = Office.MsoAutoShapeType.msoShapeRoundedRectangularCallout,
            ["OvalCallout"] = Office.MsoAutoShapeType.msoShapeOvalCallout,
            ["CloudCallout"] = Office.MsoAutoShapeType.msoShapeCloudCallout,
            ["LineCallout1"] = Office.MsoAutoShapeType.msoShapeLineCallout1,
            ["LineCallout2"] = Office.MsoAutoShapeType.msoShapeLineCallout2,
            ["LineCallout3"] = Office.MsoAutoShapeType.msoShapeLineCallout3,
            ["LineCallout4"] = Office.MsoAutoShapeType.msoShapeLineCallout4,
            ["LineCallout1AccentBar"] = Office.MsoAutoShapeType.msoShapeLineCallout1AccentBar,
            ["LineCallout2AccentBar"] = Office.MsoAutoShapeType.msoShapeLineCallout2AccentBar,
            ["LineCallout3AccentBar"] = Office.MsoAutoShapeType.msoShapeLineCallout3AccentBar,
            ["LineCallout4AccentBar"] = Office.MsoAutoShapeType.msoShapeLineCallout4AccentBar,
            ["LineCallout1NoBorder"] = Office.MsoAutoShapeType.msoShapeLineCallout1NoBorder,
            ["LineCallout2NoBorder"] = Office.MsoAutoShapeType.msoShapeLineCallout2NoBorder,
            ["LineCallout3NoBorder"] = Office.MsoAutoShapeType.msoShapeLineCallout3NoBorder,
            ["LineCallout4NoBorder"] = Office.MsoAutoShapeType.msoShapeLineCallout4NoBorder,
            ["LineCallout1BorderandAccentBar"] = Office.MsoAutoShapeType.msoShapeLineCallout1BorderandAccentBar,
            ["LineCallout2BorderandAccentBar"] = Office.MsoAutoShapeType.msoShapeLineCallout2BorderandAccentBar,
            ["LineCallout3BorderandAccentBar"] = Office.MsoAutoShapeType.msoShapeLineCallout3BorderandAccentBar,
            ["LineCallout4BorderandAccentBar"] = Office.MsoAutoShapeType.msoShapeLineCallout4BorderandAccentBar,
            ["ActionButtonCustom"] = Office.MsoAutoShapeType.msoShapeActionButtonCustom,
            ["ActionButtonHome"] = Office.MsoAutoShapeType.msoShapeActionButtonHome,
            ["ActionButtonHelp"] = Office.MsoAutoShapeType.msoShapeActionButtonHelp,
            ["ActionButtonInformation"] = Office.MsoAutoShapeType.msoShapeActionButtonInformation,
            ["ActionButtonBackorPrevious"] = Office.MsoAutoShapeType.msoShapeActionButtonBackorPrevious,
            ["ActionButtonForwardorNext"] = Office.MsoAutoShapeType.msoShapeActionButtonForwardorNext,
            ["ActionButtonBeginning"] = Office.MsoAutoShapeType.msoShapeActionButtonBeginning,
            ["ActionButtonEnd"] = Office.MsoAutoShapeType.msoShapeActionButtonEnd,
            ["ActionButtonReturn"] = Office.MsoAutoShapeType.msoShapeActionButtonReturn,
            ["ActionButtonDocument"] = Office.MsoAutoShapeType.msoShapeActionButtonDocument,
            ["ActionButtonSound"] = Office.MsoAutoShapeType.msoShapeActionButtonSound,
            ["ActionButtonMovie"] = Office.MsoAutoShapeType.msoShapeActionButtonMovie,
            ["Balloon"] = Office.MsoAutoShapeType.msoShapeBalloon,
            ["NotPrimitive"] = Office.MsoAutoShapeType.msoShapeNotPrimitive,
            ["FlowchartOfflineStorage"] = Office.MsoAutoShapeType.msoShapeFlowchartOfflineStorage,
            ["LeftRightRibbon"] = Office.MsoAutoShapeType.msoShapeLeftRightRibbon,
            ["DiagonalStripe"] = Office.MsoAutoShapeType.msoShapeDiagonalStripe,
            ["Pie"] = Office.MsoAutoShapeType.msoShapePie,
            ["NonIsoscelesTrapezoid"] = Office.MsoAutoShapeType.msoShapeNonIsoscelesTrapezoid,
            ["Decagon"] = Office.MsoAutoShapeType.msoShapeDecagon,
            ["Heptagon"] = Office.MsoAutoShapeType.msoShapeHeptagon,
            ["Dodecagon"] = Office.MsoAutoShapeType.msoShapeDodecagon,
            ["6pointStar"] = Office.MsoAutoShapeType.msoShape6pointStar,
            ["7pointStar"] = Office.MsoAutoShapeType.msoShape7pointStar,
            ["10pointStar"] = Office.MsoAutoShapeType.msoShape10pointStar,
            ["12pointStar"] = Office.MsoAutoShapeType.msoShape12pointStar,
            ["Round1Rectangle"] = Office.MsoAutoShapeType.msoShapeRound1Rectangle,
            ["Round2SameRectangle"] = Office.MsoAutoShapeType.msoShapeRound2SameRectangle,
            ["Round2DiagRectangle"] = Office.MsoAutoShapeType.msoShapeRound2DiagRectangle,
            ["SnipRoundRectangle"] = Office.MsoAutoShapeType.msoShapeSnipRoundRectangle,
            ["Snip1Rectangle"] = Office.MsoAutoShapeType.msoShapeSnip1Rectangle,
            ["Snip2SameRectangle"] = Office.MsoAutoShapeType.msoShapeSnip2SameRectangle,
            ["Snip2DiagRectangle"] = Office.MsoAutoShapeType.msoShapeSnip2DiagRectangle,
            ["Frame"] = Office.MsoAutoShapeType.msoShapeFrame,
            ["HalfFrame"] = Office.MsoAutoShapeType.msoShapeHalfFrame,
            ["Tear"] = Office.MsoAutoShapeType.msoShapeTear,
            ["Chord"] = Office.MsoAutoShapeType.msoShapeChord,
            ["Corner"] = Office.MsoAutoShapeType.msoShapeCorner,
            ["MathPlus"] = Office.MsoAutoShapeType.msoShapeMathPlus,
            ["MathMinus"] = Office.MsoAutoShapeType.msoShapeMathMinus,
            ["MathMultiply"] = Office.MsoAutoShapeType.msoShapeMathMultiply,
            ["MathDivide"] = Office.MsoAutoShapeType.msoShapeMathDivide,
            ["MathEqual"] = Office.MsoAutoShapeType.msoShapeMathEqual,
            ["MathNotEqual"] = Office.MsoAutoShapeType.msoShapeMathNotEqual,
            ["CornerTabs"] = Office.MsoAutoShapeType.msoShapeCornerTabs,
            ["SquareTabs"] = Office.MsoAutoShapeType.msoShapeSquareTabs,
            ["PlaqueTabs"] = Office.MsoAutoShapeType.msoShapePlaqueTabs,
            ["Gear6"] = Office.MsoAutoShapeType.msoShapeGear6,
            ["Gear9"] = Office.MsoAutoShapeType.msoShapeGear9,
            ["Funnel"] = Office.MsoAutoShapeType.msoShapeFunnel,
            ["PieWedge"] = Office.MsoAutoShapeType.msoShapePieWedge,
            ["LeftCircularArrow"] = Office.MsoAutoShapeType.msoShapeLeftCircularArrow,
            ["LeftRightCircularArrow"] = Office.MsoAutoShapeType.msoShapeLeftRightCircularArrow,
            ["SwooshArrow"] = Office.MsoAutoShapeType.msoShapeSwooshArrow,
            ["Cloud"] = Office.MsoAutoShapeType.msoShapeCloud,
            ["ChartX"] = Office.MsoAutoShapeType.msoShapeChartX,
            ["ChartStar"] = Office.MsoAutoShapeType.msoShapeChartStar,
            ["ChartPlus"] = Office.MsoAutoShapeType.msoShapeChartPlus,
            ["LineInverse"] = Office.MsoAutoShapeType.msoShapeLineInverse,

        }.ToImmutableDictionary();

        public static readonly ImmutableDictionary<string, string> MSO_IMAGES = new Dictionary<string, string>()
        {
            ["Rectangle"] = "ShapeRectangle",
            ["RoundedRectangle"] = "ShapeRoundedRectangle",
            ["IsoscelesTriangle"] = "ShapeIsoscelesTriangle",
            ["Oval"] = "ShapeOval",
            ["SmileyFace"] = "ShapeSmileyFace",
            ["Donut"] = "ShapeDonut",
            ["Heart"] = "ShapeHeart",
            ["RightArrow"] = "ShapeRightArrow",
            ["LeftArrow"] = "ShapeLeftArrow",
            ["UpArrow"] = "ShapeUpArrow",
            ["DownArrow"] = "ShapeDownArrow",
            ["RoundedRectangularCallout"] = "ShapeRoundedRectangularCallout",
            ["5pointStar"] = "ShapeStar",
            ["8pointStar"] = "ShapeSeal8",
            ["16pointStar"] = "ShapeSeal16",
            ["24pointStar"] = "ShapeSeal24",
        }.ToImmutableDictionary();

        public ChangeAutoShapeType() : base("change_shape_type")
        {
        }

        private Office.MsoAutoShapeType? GetShapeType(string arg)
        {
            if (SHAPES.TryGetValue(arg, out Office.MsoAutoShapeType result))
            {
                return result;
            }
            return null;
        }

        protected override IEnumerable<PowerPoint.Shape> GetSelectedShapes()
        {
            return from shape in base.GetSelectedShapes()
                   where shape.Type == Office.MsoShapeType.msoAutoShape
                   where shape.AutoShapeType != Office.MsoAutoShapeType.msoShapeMixed
                   select shape;
        }

        public override Bitmap? GetImage(string controlId, string arg = "")
        {
            var result = Utils.LoadImageResource($"hztoolbar.icons.{controlId}.png");
            return result;
        }

        public override string? GetMsoImage(string controlId, string arg = "")
        {
            if (MSO_IMAGES.TryGetValue(arg, out var result))
            {
                return result;
            }
            return null;
        }

        public override bool IsEnabled(string arg = "")
        {
            if (GetShapeType(arg) == null)
            {
                return false;
            }
            return GetSelectedShapes().Take(1).Count() > 0;
        }

        public override bool Run(string arg = "")
        {
            var shapeType = GetShapeType(arg);
            if (shapeType != null)
            {
                foreach (var shape in GetSelectedShapes())
                {
                    shape.AutoShapeType = shapeType.Value;
                }
            }
            Properties.Settings.Default.change_shape_type = arg;
            Properties.Settings.Default.Save();
            return true;
        }
    }


    public class RepeatLastChangeAction : ToolbarAction
    {

        private readonly ChangeAutoShapeType changeAction = new ChangeAutoShapeType();

        public RepeatLastChangeAction() : base("repeat_last_shape_change")
        {

        }

        public override Bitmap? GetImage(string controlId, string arg = "")
        {
            return changeAction.GetImage(controlId, Properties.Settings.Default.change_shape_type);
        }

        public override string? GetMsoImage(string controlId, string arg = "")
        {
            return changeAction.GetMsoImage(controlId, Properties.Settings.Default.change_shape_type);
        }

        public override bool IsEnabled(string arg = "")
        {
            return changeAction.IsEnabled(Properties.Settings.Default.change_shape_type);
        }

        public override bool Run(string arg = "")
        {
            changeAction.Run(Properties.Settings.Default.change_shape_type);
            return false;
        }
    }


    public class CopyShapeTypeAction : ToolbarAction
    {
        public CopyShapeTypeAction() : base("copy_shape_type") { }

        protected override IEnumerable<PowerPoint.Shape> GetSelectedShapes()
        {
            return (
                from shape in base.GetSelectedShapes()
                where shape.Type == Office.MsoShapeType.msoAutoShape
                where shape.AutoShapeType != Office.MsoAutoShapeType.msoShapeMixed
                select shape
            ).ToList();
        }

        public override bool IsEnabled(string arg = "")
        {
            return GetSelectedShapes().Take(2).Count() > 0;
        }

        public override bool Run(string arg = "")
        {
            var shapes = GetSelectedShapes().ToList();
            if (shapes.Count > 1) {
                var reference = shapes[0];
                foreach (var shape in shapes)
                {
                    if (shape == reference) {
                        continue;
                    }
                    shape.AutoShapeType = reference.AutoShapeType;
                }
            }
            return true;
        }
    }
}
