using System.Collections.Generic;
using System.Globalization;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using OpenXMLTemplates.Variables;
using OpenXMLTemplates.Variables.Exceptions;

namespace OpenXMLTempaltesTest
{
    public class DataTests
    {
        [Test]
        public void GetVariableByIdentifierWorksCorrectly()
        {
            var addressDict = new Dictionary<string, string>
                { { "street", "MyStreet" }, { "number", "1" }, { "app", "2" } };


            var data = new Dictionary<string, object>
            {
                { "name", "MyName" }, { "address", addressDict }, { "phones", new List<string> { "123", "12345" } }
            };

            var source = new VariableSource(data);

            ClassicAssert.AreEqual("MyName", source.GetVariable<string>("name"));
            ClassicAssert.AreEqual("MyStreet", source.GetVariable<string>("address.street"));

            ClassicAssert.AreEqual("12345", source.GetVariable<string>("phones.[1]"));

            ClassicAssert.Throws<VariableNotFoundException>(() => source.GetVariable<string>("name.street"));
            ClassicAssert.Throws<VariableNotFoundException>(() => source.GetVariable<string>("address.streeets"));
            ClassicAssert.Throws<IncorrectVariableTypeException>(() => source.GetVariable<int>("name"));
        }

        [Test]
        public void Format_Numeric_Fields_Value_Null()
        {
            var data = new Dictionary<string, object>
            {
                { "prices", null }
            };

            var source = new VariableSource(data);

            ClassicAssert.AreEqual("", source.GetVariable<string>("prices(N2)"));
        }

        [Test]
        public void Format_Numeric_Fields()
        {
            // Set the current culture to invariant culture for consistent numeric formatting across different environments. 
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;

            var data = new Dictionary<string, object>
            {
                { "prices", new List<string> { "123", "12345.0001" } }
            };

            var source = new VariableSource(data);

            ClassicAssert.AreEqual("12,345.00", source.GetVariable<string>("prices.[1](N2)"));
        }
    }
}