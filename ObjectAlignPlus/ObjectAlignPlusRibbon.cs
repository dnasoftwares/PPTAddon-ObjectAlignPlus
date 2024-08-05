using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
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

        private RectangleF GetShapeExactRect(Shape shape)
        {
            //Shapeの位置、サイズ、回転から正確な外形矩形を得る
            //頂点は基準点から反時計回り
            var verts = new PointF[4];
            verts[0] = new PointF(shape.Left, shape.Top);
            verts[1] = new PointF(shape.Left + shape.Width, shape.Top);
            verts[2] = new PointF(shape.Left + shape.Width, shape.Top + shape.Height);
            verts[3] = new PointF(shape.Left, shape.Top + shape.Height);

            //Shapeの回転角度を取得 (単位は度)
            var angle = shape.Rotation;
            //色々都合が良いのでラジアンにしておく
            var rad = angle * System.Math.PI / 180;

            var rotcenter = new PointF(shape.Left + shape.Width / 2, shape.Top + shape.Height / 2);

            //0,90,180,270度の場合は三角関数を使わずに計算
            if (angle == 0 || angle==180)
            {
                return new RectangleF(verts[0], new SizeF(shape.Width, shape.Height));
            }
            else if (angle == 90 || angle==270)
            {
                return new RectangleF(new PointF(rotcenter.X-shape.Height/2,rotcenter.Y-shape.Width/2),
                    new SizeF(shape.Height, shape.Width));
            }

            //回転中心を原点にして、頂点を回転させる
            for (int i = 0; i < 4; i++)
            {
                var x = verts[i].X - rotcenter.X;
                var y = verts[i].Y - rotcenter.Y;
                verts[i].X = (float)(x * System.Math.Cos(rad) - y * System.Math.Sin(rad) + rotcenter.X);
                verts[i].Y = (float)(x * System.Math.Sin(rad) + y * System.Math.Cos(rad) + rotcenter.Y);
            }
            //頂点座標が出たらそれらの座標から外接矩形を求める
            var left = verts.Select(v => v.X).Min();
            var top = verts.Select(v => v.Y).Min();
            var right = verts.Select(v => v.X).Max();
            var bottom = verts.Select(v => v.Y).Max();
            return new RectangleF(left, top, right - left, bottom - top);
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
            shapeList.Sort((a, b) =>
                GetShapeExactRect(a).Left.CompareTo(GetShapeExactRect(b).Left));

            var rect0 = GetShapeExactRect(shapeList[0]);

            //左端＝[0]のオブジェクトは基準なので動かさない。[0]の右端＋spacingからはじめる
            var x = rect0.Right + spacing;
            for(int i = 1; i < shapeList.Count; i++)
            {
                var rect = GetShapeExactRect(shapeList[i]);
                //ShapeのLeftと外接矩形のLeftが違うので、外接矩形のLeft同士で差分を取って平行移動
                shapeList[i].IncrementLeft(x-rect.Left);
                x += rect.Width + spacing;
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
            shapeList.Sort((a, b) =>
                GetShapeExactRect(b).Right.CompareTo(GetShapeExactRect(a).Right)
            );

            //右端＝[0]のオブジェクトは基準なので動かさない。[0]の左端＋spacingからはじめる
            var rect0 = GetShapeExactRect(shapeList[0]);

            var x = rect0.Left - spacing;
            for (int i = 1; i < shapeList.Count; i++)
            {
                var rect = GetShapeExactRect(shapeList[i]);
                shapeList[i].IncrementLeft(x - rect.Right);
                x -= rect.Width + spacing;
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
            shapeList.Sort((a, b) => GetShapeExactRect(a).Top.CompareTo(GetShapeExactRect(b).Top));

            //上端＝[0]のオブジェクトは基準なので動かさない。[0]の下端+spacingからはじめる
            var y = GetShapeExactRect(shapeList[0]).Bottom + spacing;
            for (int i = 1; i < shapeList.Count; i++)
            {
                var rect = GetShapeExactRect(shapeList[i]);
                shapeList[i].IncrementTop(y- rect.Top);
                y += rect.Height + spacing;
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
            shapeList.Sort((a, b) => GetShapeExactRect(b).Bottom.CompareTo(GetShapeExactRect(a).Bottom));

            //下端＝[0]のオブジェクトは基準なので動かさない。[0]の上端-spacingからはじめる
            var rect0 = GetShapeExactRect(shapeList[0]);
            var y = rect0.Top - spacing;
            for (int i = 1; i < shapeList.Count; i++)
            {
                var rect = GetShapeExactRect(shapeList[i]);
                shapeList[i].IncrementTop(y - rect.Bottom);
                y -= rect.Height + spacing;
            }
        }
    }
}
