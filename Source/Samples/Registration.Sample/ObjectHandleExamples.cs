using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace Registration.Sample
{
    [ExcelMarshalByHandle]
    public interface IMarshaledByHandle
    {
        double Property { get; }
    }

    [ExcelMarshalByHandle]
    public class SampleClass1 : IMarshaledByHandle
    {
        public double Property { private set; get; }
        public SampleClass1(double p)
        {
            Property = p;
        }
    }

    [ExcelMarshalByHandle]
    public class SampleClass2 : IMarshaledByHandle
    {
        public double Property { private set; get; }
        public SampleClass2(double p)
        {
            Property = p*p;
        }
    }

    [ExcelMarshalByHandle]
    public class Compound : IMarshaledByHandle
    {
        public double Property { private set; get; }
        public Compound(IMarshaledByHandle c1, IMarshaledByHandle c2)
        {
            Property = c1.Property + c2.Property;
        }
    }

    public enum SomeEnum
    {
        One,
        Two,
        Three
    }

    public static class ObjectHandleExamples
    {
        [ExcelMapArrayFunction]
        public static IEnumerable<IMarshaledByHandle> dnaFactoryMultiple(IEnumerable<SomeEnum> enumValues,
            IEnumerable<double> doubleValues)
        {
            var enumsIter = enumValues.GetEnumerator();
            var valuesIter = doubleValues.GetEnumerator();
            while(enumsIter.MoveNext() && valuesIter.MoveNext())
            {
                yield return dnaFactorySingle(enumsIter.Current, valuesIter.Current);
            }
        }

        [ExcelFunction]
        public static IMarshaledByHandle dnaFactorySingle(SomeEnum enumValue, double doubleValue)
        {
            IMarshaledByHandle item;
            switch (enumValue)
            {
                case SomeEnum.One:
                    item = new SampleClass1(doubleValue);
                    break;
                case SomeEnum.Two:
                    item = new SampleClass2(doubleValue);
                    break;
                default:
                    throw new ArgumentException($"Don't know how to create an object of type {enumValue}.");
            }
            return item;
        }

        [ExcelMapArrayFunction]
        public static double dnaUseSomeHandles(IEnumerable<IMarshaledByHandle> objects)
        {
            return objects.Sum(x => x.Property);
        }

        [ExcelFunction]
        public static IMarshaledByHandle dnaFactoryCompound(IMarshaledByHandle c1, IMarshaledByHandle c2)
        {
            return new Compound(c1, c2);
        }
    }
}
