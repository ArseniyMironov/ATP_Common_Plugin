using ATP_Common_Plugin.Utils.Spaces;
using Autodesk.Revit.DB;
using System;

namespace ATP_Common_Plugin.Utils.Geometry
{
    public static class TransformUtils
    {
        private const double Eps = 1e-9;

        public static bool IsFaceHorizontal(Face face)
        {
            PlanarFace pf = face as PlanarFace;
            if (pf == null) return false;
            XYZ n = pf.FaceNormal;
            return Math.Abs(n.Z) > 0.999;
        }

        public static bool IsZeroLength(this XYZ v)
        {
            return v == null || (Math.Abs(v.X) < Eps && Math.Abs(v.Y) < Eps && System.Math.Abs(v.Z) < Eps);
        }

        public static Models.Spaces.Orientation4 QuantizeTo4(double angleRadians)
        {
            // 0 = East; convert so 0 = North
            double a = angleRadians - Math.PI / 2.0;
            while (a < 0) a += 2 * Math.PI;
            while (a >= 2 * Math.PI) a -= 2 * Math.PI;

            // sectors: N(315..45), E(45..135), S(135..225), W(225..315)
            double deg = a * 180.0 / Math.PI;
            if (deg >= 315 || deg < 45) return Models.Spaces.Orientation4.N;
            if (deg >= 45 && deg < 135) return Models.Spaces.Orientation4.E;
            if (deg >= 135 && deg < 225) return Models.Spaces.Orientation4.S;
            return Models.Spaces.Orientation4.W;
        }

        /// <summary>
        /// Определяет, какая из двух сторон нормали hostFace соответствует
        /// «внутрь Space» (nin) и «наружу от Space» (nout).
        /// </summary>
        public static void SelectSpaceOrientedNormals(
            SpatialElement space,
            XYZ pointOnHostFace,
            XYZ hostFaceNormal,
            out XYZ nin, out XYZ nout)
        {
            // пробуем шагом EPS понять, какая сторона попадает в Space
            double eps = UnitUtilsEx.MetersToFeet(Models.Settings.Epsilon);
            XYZ pInCandidate = pointOnHostFace - hostFaceNormal.Multiply(eps);
            bool insideMinus = SpacePointTests.IsPointInSpaceFast(space, pInCandidate);

            if (insideMinus)
            {
                // -n → внутрь помещения
                nin = hostFaceNormal.Negate();
                nout = hostFaceNormal;
            }
            else
            {
                // проверяем другую сторону
                XYZ pInOther = pointOnHostFace + hostFaceNormal.Multiply(eps);
                bool insidePlus = SpacePointTests.IsPointInSpaceFast(space, pInOther);

                if (insidePlus)
                {
                    nin = hostFaceNormal;
                    nout = hostFaceNormal.Negate();
                }
                else
                {
                    // Ни одна сторона не распознана однозначно (краевой случай) —
                    // дефолтимся: считаем +n наружу.
                    nin = hostFaceNormal.Negate();
                    nout = hostFaceNormal;
                }
            }
        }
    }
}