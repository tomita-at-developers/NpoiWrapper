using System;
using System.Linq.Expressions;

namespace Developers.NpoiWrapper.Utils
{
    public static class NameOf<TSource>
    {
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

        public static string FullName(string sourceFieldName, Expression<Func<TSource, object>> expression)
        {
            var result = FullName(expression);
            result = string.IsNullOrEmpty(sourceFieldName) ? result : sourceFieldName + "." + result;
            return result;
        }
    }
}
