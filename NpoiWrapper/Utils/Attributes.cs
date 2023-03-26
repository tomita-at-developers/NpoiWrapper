using System;
using System.Collections.Generic;
using System.Reflection;

namespace Developers.NpoiWrapper.Utils
{
    /// <summary>
    /// カスタムプロパティ：Import
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    internal class ImportAttribute : Attribute
    {
        public bool Import { get; set; }

        public ImportAttribute(bool Import)
        {
            this.Import = Import;
        }
    }

    /// <summary>
    /// カスタムプロパティ：Compare
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    internal class ComparisonAttribute : Attribute
    {
        public bool Compare { get; set; }

        public ComparisonAttribute(bool Comparison)
        {
            this.Compare = Comparison;
        }
    }

    /// <summary>
    /// カスタムプロパティ：Export
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    internal class ExportAttribute : Attribute
    {
        public bool Export { get; set; }

        public ExportAttribute(bool Export)
        {
            this.Export = Export;
        }
    }

    /// <summary>
    /// AttributeUtilityクラス
    /// </summary>
    internal static class AttributeUtility
    {
        //
        // T で指定したプロパティを1つだけ取得
        //

        // 型を指定して Public プロパティの属性を取得する
        public static T GetPropertyAttribute<T>(Type type, string name) where T : Attribute
        {
            var prop = type.GetProperty(name);
            if (prop == null)
            {
                // 指定したプロパティが見つからない
                return default;
            }
            var att = prop.GetCustomAttribute<T>();
            if (att == null)
            {
                // 指定した属性が付与されていない
                return default;
            }
            return att;
        }

        /// <summary>
        /// インスタンスを指定してプロパティの属性を取得 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="instance"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static T GetPropertyAttribute<T>(object instance, string name) where T : Attribute
        {
            return GetPropertyAttribute<T>(instance.GetType(), name);
        }

        /// <summary>
        /// 型を指定してプロパティに付与されているすべてのプロパティを取得する 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static IEnumerable<Attribute> GetPropertyAttributes(Type type, string name)
        {
            var prop = type.GetProperty(name, BindingFlags.Public | BindingFlags.NonPublic);
            if (prop == null)
            {
                // 指定したプロパティが見つからない
                return default;
            }

            return prop.GetCustomAttributes<Attribute>();
        }

        // インスタンスを指定してプロパティに付与されているすべてのプロパティを取得
        public static IEnumerable<Attribute> GetPropertyAttributes(object instance, string name)
        {
            return GetPropertyAttributes(instance.GetType(), name);
        }
    }
}
