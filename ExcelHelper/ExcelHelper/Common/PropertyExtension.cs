﻿using System;
using System.Reflection;

namespace ExcelHelper.Common
{
    public static class PropertyExtension
    {
        public static bool IsNullable(this PropertyInfo property)
        {
            return property.PropertyType.IsGenericType && property.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>);
        }
    }
}
