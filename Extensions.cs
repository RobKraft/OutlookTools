using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OutlookTools
{
	public static class Extensions
	{
		//https://www.c-sharpcorner.com/article/dynamic-sorting-orderby-based-on-user-preference/

		public static IQueryable<T> OrderByMe<T>(this IQueryable<T> source, string columnName, bool isAscending = true)
		{
			if (String.IsNullOrEmpty(columnName))
			{
				return source;
			}

			ParameterExpression parameter = Expression.Parameter(source.ElementType, "");

			MemberExpression property = Expression.Property(parameter, columnName);
			LambdaExpression lambda = Expression.Lambda(property, parameter);

			string methodName = isAscending ? "OrderBy" : "OrderByDescending";

			Expression methodCallExpression = Expression.Call(typeof(Queryable), methodName,
								  new Type[] { source.ElementType, property.Type },
								  source.Expression, Expression.Quote(lambda));

			return source.Provider.CreateQuery<T>(methodCallExpression);
		}
	}
}
