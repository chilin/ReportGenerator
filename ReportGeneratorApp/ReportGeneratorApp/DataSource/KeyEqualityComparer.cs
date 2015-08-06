using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
///KeyEqualityComparer 的摘要说明
/// </summary>
public class KeyEqualityComparer<T> : IEqualityComparer<T>
{
    private readonly Func<T, object> keyExtractor;

    public KeyEqualityComparer(Func<T, object> keyExtractor)
    {
        this.keyExtractor = keyExtractor;
    }

    public bool Equals(T x, T y)
    {
        return this.keyExtractor(x).Equals(this.keyExtractor(y));
    }

    public int GetHashCode(T obj)
    {
        return this.keyExtractor(obj).GetHashCode();
    }
}