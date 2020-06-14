using System.Collections.Generic;
using NUnit.Framework;
using OpenXMLTemplates.Variables;
using OpenXMLTemplates.Variables.Exceptions;

namespace OpenXMLTempaltesTest
{
    public class DataTests
    {
        
        [Test]
        public void GetVariableByIdentifierWorksCorrectly()
        {
            var addressDict = new Dictionary<string, string> {{"street", "MyStreet"}, {"number", "1"}, {"app", "2"}};


            var data = new Dictionary<string, object>
            {
                {"name", "MyName"}, {"address", addressDict}, {"phones", new List<string> {"123", "12345"}}
            };

            var source = new VariableSource(data);
            
            Assert.AreEqual("MyName", source.GetVariable<string>("name"));
            Assert.AreEqual("MyStreet", source.GetVariable<string>("address.street"));

            Assert.AreEqual("12345", source.GetVariable<string>("phones.[1]"));

            Assert.Throws<VariableNotFoundException>((() => source.GetVariable<string>("address.streeets")));
            Assert.Throws<IncorrectIdentifierException>((() => source.GetVariable<string>("name.street")));
            Assert.Throws<IncorrectVariableTypeException>((() => source.GetVariable<int>("name")));
        }
    }
}