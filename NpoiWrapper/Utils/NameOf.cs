using System;
using System.Linq.Expressions;

namespace Developers.NpoiWrapper.Utils
{
    /// <summary>
    /// プロパティ名取得クラス
    /// </summary>
    /// <typeparam name="TSource"></typeparam>
    internal static class NameOf<TSource>
    {
        /// <summary>
        /// プロパティ名の取得。
        /// nameofでは最後のプロパティ名しか取得できないが、このメソッドではすべて取得できる、
        /// 使い方；NameOf<クラス名>.FullName(n => n.プロパティ名.子プロパティ名,,,);
        /// </summary>
        /// <param name="expression"></param>
        /// <returns></returns>
        public static string FullName(Expression<Func<TSource, object>> expression)
        {
            var memberExpression = expression.Body as MemberExpression;
            if (memberExpression == null)
            {
                if (expression.Body is UnaryExpression unaryExpression && unaryExpression.NodeType == ExpressionType.Convert)
                    memberExpression = unaryExpression.Operand as MemberExpression;
            }

            var result = memberExpression.ToString();
            result = result.Substring(result.IndexOf('.') + 1);

            return result;
        }
        /// <summary>
        /// プロパティ名の取得。
        /// クラス名も付加したい場合に使用する。
        /// </summary>
        /// <param name="sourceFieldName">先頭に付加したいクラス名</param>
        /// <param name="expression"></param>
        /// <returns></returns>
        public static string FullName(string sourceFieldName, Expression<Func<TSource, object>> expression)
        {
            var result = FullName(expression);
            result = string.IsNullOrEmpty(sourceFieldName) ? result : sourceFieldName + "." + result;
            return result;
        }
    }
}
