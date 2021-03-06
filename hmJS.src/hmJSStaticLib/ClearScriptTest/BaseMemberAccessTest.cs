// 
// Copyright (c) Microsoft Corporation. All rights reserved.
// 
// Microsoft Public License (MS-PL)
// 
// This license governs use of the accompanying software. If you use the
// software, you accept this license. If you do not accept the license, do not
// use the software.
// 
// 1. Definitions
// 
//   The terms "reproduce," "reproduction," "derivative works," and
//   "distribution" have the same meaning here as under U.S. copyright law. A
//   "contribution" is the original software, or any additions or changes to
//   the software. A "contributor" is any person that distributes its
//   contribution under this license. "Licensed patents" are a contributor's
//   patent claims that read directly on its contribution.
// 
// 2. Grant of Rights
// 
//   (A) Copyright Grant- Subject to the terms of this license, including the
//       license conditions and limitations in section 3, each contributor
//       grants you a non-exclusive, worldwide, royalty-free copyright license
//       to reproduce its contribution, prepare derivative works of its
//       contribution, and distribute its contribution or any derivative works
//       that you create.
// 
//   (B) Patent Grant- Subject to the terms of this license, including the
//       license conditions and limitations in section 3, each contributor
//       grants you a non-exclusive, worldwide, royalty-free license under its
//       licensed patents to make, have made, use, sell, offer for sale,
//       import, and/or otherwise dispose of its contribution in the software
//       or derivative works of the contribution in the software.
// 
// 3. Conditions and Limitations
// 
//   (A) No Trademark License- This license does not grant you rights to use
//       any contributors' name, logo, or trademarks.
// 
//   (B) If you bring a patent claim against any contributor over patents that
//       you claim are infringed by the software, your patent license from such
//       contributor to the software ends automatically.
// 
//   (C) If you distribute any portion of the software, you must retain all
//       copyright, patent, trademark, and attribution notices that are present
//       in the software.
// 
//   (D) If you distribute any portion of the software in source code form, you
//       may do so only under this license by including a complete copy of this
//       license with your distribution. If you distribute any portion of the
//       software in compiled or object code form, you may only do so under a
//       license that complies with this license.
// 
//   (E) The software is licensed "as-is." You bear the risk of using it. The
//       contributors give no express warranties, guarantees or conditions. You
//       may have additional consumer rights under your local laws which this
//       license cannot change. To the extent permitted under your local laws,
//       the contributors exclude the implied warranties of merchantability,
//       fitness for a particular purpose and non-infringement.
//       

using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.ClearScript.V8;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.ClearScript.Test
{
    [TestClass]
    [DeploymentItem("ClearScriptV8-64.dll")]
    [DeploymentItem("ClearScriptV8-32.dll")]
    [DeploymentItem("v8-x64.dll")]
    [DeploymentItem("v8-ia32.dll")]
    [SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable", Justification = "Test classes use TestCleanupAttribute for deterministic teardown.")]
    public class BaseMemberAccessTest : ClearScriptTest
    {
        #region setup / teardown

        private ScriptEngine engine;
        private TestObject testObject;

        [TestInitialize]
        public void TestInitialize()
        {
            engine = new V8ScriptEngine(V8ScriptEngineFlags.EnableDebugging);
            engine.AddHostObject("host", new ExtendedHostFunctions());
            engine.AddHostObject("mscorlib", HostItemFlags.GlobalMembers, new HostTypeCollection("mscorlib"));
            engine.AddHostObject("ClearScriptTest", HostItemFlags.GlobalMembers, new HostTypeCollection("ClearScriptTest").GetNamespaceNode("Microsoft.ClearScript.Test"));
            engine.AddHostObject("testObject", testObject = new TestObject());
        }

        [TestCleanup]
        public void TestCleanup()
        {
            testObject = null;
            engine.Dispose();
            BaseTestCleanup();
        }

        #endregion

        #region test methods

        // ReSharper disable InconsistentNaming

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field()
        {
            testObject.BaseField = Enumerable.Range(0, 10).ToArray();
            Assert.AreEqual(10, engine.Evaluate("testObject.BaseField.Length"));
            engine.Execute("testObject.BaseField = host.newArr(System.Int32, 5)");
            Assert.AreEqual(5, testObject.BaseField.Length);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field_Null()
        {
            engine.Execute("testObject.BaseField = null");
            Assert.IsNull(testObject.BaseField);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field_BadAssignment()
        {
            TestUtil.AssertException<ArgumentException>(() => engine.Execute("testObject.BaseField = host.newArr(System.Double, 5)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field_Scalar()
        {
            testObject.BaseScalarField = 12345;
            Assert.AreEqual(12345, engine.Evaluate("testObject.BaseScalarField"));
            engine.Execute("testObject.BaseScalarField = 4321");
            Assert.AreEqual(4321, testObject.BaseScalarField);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field_Scalar_Overflow()
        {
            TestUtil.AssertException<OverflowException>(() => engine.Execute("testObject.BaseScalarField = 54321"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field_Scalar_BadAssignment()
        {
            TestUtil.AssertException<ArgumentException>(() => engine.Execute("testObject.BaseScalarField = TestEnum.Second"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field_Enum()
        {
            testObject.BaseEnumField = TestEnum.Second;
            Assert.AreEqual(TestEnum.Second, engine.Evaluate("testObject.BaseEnumField"));
            engine.Execute("testObject.BaseEnumField = TestEnum.Third");
            Assert.AreEqual(TestEnum.Third, testObject.BaseEnumField);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field_Enum_Zero()
        {
            engine.Execute("testObject.BaseEnumField = 0");
            Assert.AreEqual((TestEnum)0, testObject.BaseEnumField);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field_Enum_BadAssignment()
        {
            TestUtil.AssertException<ArgumentException>(() => engine.Execute("testObject.BaseEnumField = 1"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field_Struct()
        {
            testObject.BaseStructField = TimeSpan.FromDays(5);
            Assert.AreEqual(TimeSpan.FromDays(5), engine.Evaluate("testObject.BaseStructField"));
            engine.Execute("testObject.BaseStructField = System.TimeSpan.FromSeconds(25)");
            Assert.AreEqual(TimeSpan.FromSeconds(25), testObject.BaseStructField);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Field_Struct_BadAssignment()
        {
            TestUtil.AssertException<ArgumentException>(() => engine.Execute("testObject.BaseStructField = System.DateTime.Now"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property()
        {
            testObject.BaseProperty = Enumerable.Range(0, 10).ToArray();
            Assert.AreEqual(10, engine.Evaluate("testObject.BaseProperty.Length"));
            engine.Execute("testObject.BaseProperty = host.newArr(System.Int32, 5)");
            Assert.AreEqual(5, testObject.BaseProperty.Length);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property_Null()
        {
            engine.Execute("testObject.BaseProperty = null");
            Assert.IsNull(testObject.BaseProperty);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property_BadAssignment()
        {
            TestUtil.AssertException<ArgumentException>(() => engine.Execute("testObject.BaseProperty = host.newArr(System.Double, 5)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property_Scalar()
        {
            testObject.BaseScalarProperty = 12345;
            Assert.AreEqual(12345, engine.Evaluate("testObject.BaseScalarProperty"));
            engine.Execute("testObject.BaseScalarProperty = 4321");
            Assert.AreEqual(4321, testObject.BaseScalarProperty);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property_Scalar_Overflow()
        {
            TestUtil.AssertException<OverflowException>(() => engine.Execute("testObject.BaseScalarProperty = 54321"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property_Scalar_BadAssignment()
        {
            TestUtil.AssertException<ArgumentException>(() => engine.Execute("testObject.BaseScalarProperty = TestEnum.Second"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property_Enum()
        {
            testObject.BaseEnumProperty = TestEnum.Second;
            Assert.AreEqual(TestEnum.Second, engine.Evaluate("testObject.BaseEnumProperty"));
            engine.Execute("testObject.BaseEnumProperty = TestEnum.Third");
            Assert.AreEqual(TestEnum.Third, testObject.BaseEnumProperty);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property_Enum_Zero()
        {
            engine.Execute("testObject.BaseEnumProperty = 0");
            Assert.AreEqual((TestEnum)0, testObject.BaseEnumProperty);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property_Enum_BadAssignment()
        {
            TestUtil.AssertException<ArgumentException>(() => engine.Execute("testObject.BaseEnumProperty = 1"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property_Struct()
        {
            testObject.BaseStructProperty = TimeSpan.FromDays(5);
            Assert.AreEqual(TimeSpan.FromDays(5), engine.Evaluate("testObject.BaseStructProperty"));
            engine.Execute("testObject.BaseStructProperty = System.TimeSpan.FromSeconds(25)");
            Assert.AreEqual(TimeSpan.FromSeconds(25), testObject.BaseStructProperty);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Property_Struct_BadAssignment()
        {
            TestUtil.AssertException<ArgumentException>(() => engine.Execute("testObject.BaseStructProperty = System.DateTime.Now"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ReadOnlyProperty()
        {
            Assert.AreEqual(testObject.BaseReadOnlyProperty, (int)engine.Evaluate("testObject.BaseReadOnlyProperty"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ReadOnlyProperty_Write()
        {
            TestUtil.AssertException<UnauthorizedAccessException>(() => engine.Execute("testObject.BaseReadOnlyProperty = 2"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Event()
        {
            engine.Execute("var connection = testObject.BaseEvent.connect(function (sender, args) { sender.BaseScalarProperty = args.Arg; })");
            testObject.BaseFireEvent(5432);
            Assert.AreEqual(5432, testObject.BaseScalarProperty);
            engine.Execute("connection.disconnect()");
            testObject.BaseFireEvent(2345);
            Assert.AreEqual(5432, testObject.BaseScalarProperty);
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Method()
        {
            Assert.AreEqual(testObject.BaseMethod("foo", 4), engine.Evaluate("testObject.BaseMethod('foo', 4)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Method_NoMatchingOverload()
        {
            TestUtil.AssertException<RuntimeBinderException>(() => engine.Execute("testObject.BaseMethod('foo', TestEnum.Second)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Method_Generic()
        {
            Assert.AreEqual(testObject.BaseMethod("foo", 4, TestEnum.Second), engine.Evaluate("testObject.BaseMethod('foo', 4, TestEnum.Second)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Method_Generic_TypeArgConstraintFailure()
        {
            TestUtil.AssertException<RuntimeBinderException>(() => engine.Execute("testObject.BaseMethod('foo', 4, testObject)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Method_GenericRedundant()
        {
            Assert.AreEqual(testObject.BaseMethod("foo", 4, TestEnum.Second), engine.Evaluate("testObject.BaseMethod(TestEnum, 'foo', 4, TestEnum.Second)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Method_GenericRedundant_MismatchedTypeArg()
        {
            TestUtil.AssertException<RuntimeBinderException>(() => engine.Execute("testObject.BaseMethod(System.Int32, 'foo', 4, TestEnum.Second)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Method_GenericExplicit()
        {
            Assert.AreEqual(testObject.BaseMethod<TestEnum>(4), engine.Evaluate("testObject.BaseMethod(TestEnum, 4)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_Method_GenericExplicit_MissingTypeArg()
        {
            TestUtil.AssertException<RuntimeBinderException>(() => engine.Execute("testObject.BaseMethod(4)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_BindTestMethod_BaseClass()
        {
            var arg = new TestArg() as BaseTestArg;
            engine.AddRestrictedHostObject("arg", arg);
            Assert.AreEqual(testObject.BaseBindTestMethod(arg), engine.Evaluate("testObject.BaseBindTestMethod(arg)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_BindTestMethod_Interface()
        {
            var arg = new TestArg() as ITestArg;
            engine.AddRestrictedHostObject("arg", arg);
            Assert.AreEqual(testObject.BaseBindTestMethod(arg), engine.Evaluate("testObject.BaseBindTestMethod(arg)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ExtensionMethod()
        {
            Assert.AreEqual(testObject.BaseExtensionMethod("foo", 4), engine.Evaluate("testObject.BaseExtensionMethod('foo', 4)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ExtensionMethod_NoMatchingOverload()
        {
            TestUtil.AssertException<RuntimeBinderException>(() => engine.Execute("testObject.BaseExtensionMethod('foo', TestEnum.Second)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ExtensionMethod_Generic()
        {
            Assert.AreEqual(testObject.BaseExtensionMethod("foo", 4, TestEnum.Second), engine.Evaluate("testObject.BaseExtensionMethod('foo', 4, TestEnum.Second)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ExtensionMethod_Generic_TypeArgConstraintFailure()
        {
            TestUtil.AssertException<RuntimeBinderException>(() => engine.Execute("testObject.BaseExtensionMethod('foo', 4, testObject)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ExtensionMethod_GenericRedundant()
        {
            Assert.AreEqual(testObject.BaseExtensionMethod("foo", 4, TestEnum.Second), engine.Evaluate("testObject.BaseExtensionMethod(TestEnum, 'foo', 4, TestEnum.Second)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ExtensionMethod_GenericRedundant_MismatchedTypeArg()
        {
            TestUtil.AssertException<RuntimeBinderException>(() => engine.Execute("testObject.BaseExtensionMethod(System.Int32, 'foo', 4, TestEnum.Second)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ExtensionMethod_GenericExplicit()
        {
            Assert.AreEqual(testObject.BaseExtensionMethod<TestEnum>(4), engine.Evaluate("testObject.BaseExtensionMethod(TestEnum, 4)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ExtensionMethod_GenericExplicit_MissingTypeArg()
        {
            TestUtil.AssertException<RuntimeBinderException>(() => engine.Execute("testObject.BaseExtensionMethod(4)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ExtensionBindTestMethod_BaseClass()
        {
            var arg = new TestArg() as BaseTestArg;
            engine.AddRestrictedHostObject("arg", arg);
            Assert.AreEqual(testObject.BaseExtensionBindTestMethod(arg), engine.Evaluate("testObject.BaseExtensionBindTestMethod(arg)"));
        }

        [TestMethod, TestCategory("BaseMemberAccess")]
        public void BaseMemberAccess_ExtensionBindTestMethod_Interface()
        {
            var arg = new TestArg() as ITestArg;
            engine.AddRestrictedHostObject("arg", arg);
            Assert.AreEqual(testObject.BaseExtensionBindTestMethod(arg), engine.Evaluate("testObject.BaseExtensionBindTestMethod(arg)"));
        }

        // ReSharper restore InconsistentNaming

        #endregion
    }
}
