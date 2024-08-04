using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;

namespace ObjectAlignPlus
{
    public partial class ObjectAlignPlusRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            ebSpacing.Text= Properties.Settings.Default.Spacing.ToString();
        }

        private void btSpPlus_Click(object sender, RibbonControlEventArgs e)
        {
            validateSpacingText();
            int spacing = int.Parse(ebSpacing.Text);
            spacing++;
            ebSpacing.Text = spacing.ToString();
        }

        private int validateSpacingText()
        {
            var spacing=0;
            if (int.TryParse(ebSpacing.Text, out spacing))
            {
                if (spacing < 0)
                {
                    ebSpacing.Text = "0";
                    spacing = 0;
                    Properties.Settings.Default.Spacing = spacing;
                }
            }
            else
            {
                ebSpacing.Text = "0";
            }
            return spacing;
        }

        private void btSpMinus_Click(object sender, RibbonControlEventArgs e)
        {
            validateSpacingText();
            int spacing = int.Parse(ebSpacing.Text);
            spacing--;
            if(spacing<0)
            {
                spacing = 0;
            }
            ebSpacing.Text = spacing.ToString();
        }

        private List<Shape> GetSelectedShape()
        {
            //選択状態のオブジェクトを取得
            var shapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            if (shapes.Count < 2)
            {
                return null;
            }

            //shapesの内容をListにコピー
            var shapeList = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
            {
                shapeList.Add(shape);
            }

            return shapeList;
        }
        private void btHorz_Click(object sender, RibbonControlEventArgs e)
        {
            var shapeList = GetSelectedShape();
            if(shapeList == null)
            {
                return;
            }

            int spacing = validateSpacingText();

            //X座標で左から順にソート
            shapeList.Sort((a, b) => a.Left.CompareTo(b.Left));

            //左端＝[0]のオブジェクトは基準なので動かさない。[0]の左端＋[0]の幅＋spacingからはじめる
            var x = shapeList[0].Left + shapeList[0].Width + spacing;
            for(int i = 1; i < shapeList.Count; i++)
            {
                shapeList[i].Left = x;
                x += shapeList[i].Width + spacing;
            }
        }


        private void btHorzRightAlign_Click(object sender, RibbonControlEventArgs e)
        {
            var shapeList = GetSelectedShape();
            if (shapeList == null)
            {
                return;
            }

            int spacing = validateSpacingText();

            //X座標で右から順にソート
            shapeList.Sort((a, b) => (b.Left+b.Width).CompareTo(a.Left+a.Width));

            //右端＝[0]のオブジェクトは基準なので動かさない。[0]の右端-spacingからはじめる
            var x = shapeList[0].Left;
            for (int i = 1; i < shapeList.Count; i++)
            {
                x -= shapeList[i].Width + spacing;
                shapeList[i].Left = x;
            }
        }

        private void btVert_Click(object sender, RibbonControlEventArgs e)
        {
            var shapeList = GetSelectedShape();
            if (shapeList == null)
            {
                return;
            }
            int spacing = validateSpacingText();

            //Y座標で上から順にソート
            shapeList.Sort((a, b) => a.Top.CompareTo(b.Top));
            var x = shapeList[0].Top+shapeList[0].Height+spacing;
            //上端＝[0]のオブジェクトは基準なので動かさない。[0]の上端＋[0]の高さ＋spacingからはじめる
            for (int i = 1; i < shapeList.Count; i++)
            {
                shapeList[i].Top = x;
                x += shapeList[i].Height + spacing;
            }
        }

        private void btVertBottomAlign_Click(object sender, RibbonControlEventArgs e)
        {
            var shapeList = GetSelectedShape();
            if (shapeList == null)
            {
                return;
            }
            int spacing = validateSpacingText();

            //Y座標で下から順にソート
            shapeList.Sort((a, b) => (b.Top + b.Height).CompareTo(a.Top + a.Height));

            //下端＝[0]のオブジェクトは基準なので動かさない。[0]の下端-spacingからはじめる
            var x = shapeList[0].Top;
            for (int i = 1; i < shapeList.Count; i++)
            {
                x -= shapeList[i].Height + spacing;
                shapeList[i].Top = x;
            }
        }
    }
}
