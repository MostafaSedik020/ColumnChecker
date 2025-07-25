using ColumnChecker.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ColumnChecker.Utils
{
    public static class MultiUtils
    {
        public static bool AreListsSimilar(List<Column> list1, List<Column> list2)
        {
            if (list1.Count != list2.Count)
                return false;

            for (int i = 0; i < list1.Count; i++)
            {
                if (!list1[i].IsSimilar(list2[i]))
                    return false;
            }

            return true;
        }
        public static double NormalizeAngleTo90(double angle)
        {
            // Reduce the angle modulo 360 to get its equivalent within [0, 360]
            double modAngle = angle % 360;
            if (modAngle < 0)
            {
                modAngle += 360; // Ensure positive angles
            }

            // Map the angle to [0, 90] using symmetry
            return modAngle % 90;
        }
    }
}
