using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelSharp.NPOI;

internal class HeaderProvider<TModel, TValue> where TModel : IFetchModel<TValue>, new()
{
    public delegate TModel GetHeaderDelagate(string[] rowNames, string[] colNames);
    public GetHeaderDelagate GetHeader { get; }

    public HeaderProvider()
    {
        var rowNames = Expression.Parameter(typeof(string[]));
        var colNames = Expression.Parameter(typeof(string[]));

        var codes = new List<Expression>();
        var header = Expression.Variable(typeof(TModel));

        codes.Add(Expression.Assign(header, Expression.New(typeof(TModel).GetConstructor([])!)));

        var props = typeof(TModel).GetProperties();
        foreach (var prop in props)
        {
            var row = prop.GetCustomAttribute<RowHeaderAttribute>();
            if (row is not null)
            {
                var setMethod = prop.GetSetMethod();
                if (setMethod is null) continue;

                var rowIndex = row.Index;
                codes.Add(
                    Expression.Call(
                        header,
                        setMethod,
                        Expression.ArrayAccess(rowNames, Expression.Constant(rowIndex, typeof(int)))
                    )
                );
                continue;
            }

            var col = prop.GetCustomAttribute<ColumnHeaderAttribute>();
            if (col is not null)
            {
                var setMethod = prop.GetSetMethod();
                if (setMethod is null) continue;

                var colIndex = col.Index;
                codes.Add(
                    Expression.Call(
                        header,
                        setMethod,
                        Expression.ArrayAccess(colNames, Expression.Constant(colIndex, typeof(int)))
                    )
                );

                continue;
            }
        }
        codes.Add(header);

        var body = Expression.Block([header], codes);
        var lambda = Expression.Lambda<GetHeaderDelagate>(body, [rowNames, colNames]);
        GetHeader = lambda.Compile();
    }
}
