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
using System.Threading;
using Microsoft.ClearScript.Util;
using Microsoft.ClearScript.V8;
using Microsoft.ClearScript.Windows;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Microsoft.ClearScript.Test
{
    [TestClass]
    [DeploymentItem("ClearScriptV8-64.dll")]
    [DeploymentItem("ClearScriptV8-32.dll")]
    [DeploymentItem("v8-x64.dll")]
    [DeploymentItem("v8-ia32.dll")]
    [SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable", Justification = "Test classes use TestCleanupAttribute for deterministic teardown.")]
    public class PropertyBagTest : ClearScriptTest
    {
        #region setup / teardown

        private ScriptEngine engine;

        [TestInitialize]
        public void TestInitialize()
        {
            engine = new JScriptEngine(WindowsScriptEngineFlags.EnableDebugging);
        }

        [TestCleanup]
        public void TestCleanup()
        {
            engine.Dispose();
            BaseTestCleanup();
        }

        #endregion

        #region test methods

        // ReSharper disable InconsistentNaming

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_Property()
        {
            var host = new HostFunctions();
            var bag = new PropertyBag { { "host", host } };
            engine.AddHostObject("bag", bag);
            Assert.AreSame(host, engine.Evaluate("bag.host"));
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_Property_Scalar()
        {
            const int value = 123;
            var bag = new PropertyBag { { "value", value } };
            engine.AddHostObject("bag", bag);
            Assert.AreEqual(value, engine.Evaluate("bag.value"));
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_Property_Struct()
        {
            var date = new DateTime(2007, 5, 22, 6, 15, 43);
            var bag = new PropertyBag { { "date", date } };
            engine.AddHostObject("bag", bag);
            Assert.AreEqual(date, engine.Evaluate("bag.date"));
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_Property_GlobalMembers()
        {
            var host = new HostFunctions();
            var bag = new PropertyBag { { "host", host } };
            engine.AddHostObject("bag", HostItemFlags.GlobalMembers, bag);
            Assert.AreSame(host, engine.Evaluate("host"));
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_HostDelegate()
        {
            var methodInvoked = false;
            Action method = () => methodInvoked = true;
            var bag = new PropertyBag { { "method", method } };
            engine.AddHostObject("bag", bag);
            engine.Execute("bag.method()");
            Assert.IsTrue(methodInvoked);
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_ScriptProperty()
        {
            engine.Execute("foo = { bar: false }");
            var bag = new PropertyBag { { "foo", engine.Script.foo } };
            engine.AddHostObject("bag", bag);
            engine.Execute("bag.foo.bar = true");
            Assert.IsTrue(engine.Script.foo.bar);
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_ScriptMethod()
        {
            engine.Execute("methodInvoked = false; function method() { methodInvoked = true }");
            var bag = new PropertyBag { { "method", engine.Script.method } };
            engine.AddHostObject("bag", bag);
            engine.Execute("bag.method()");
            Assert.IsTrue(engine.Script.methodInvoked);
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_IScriptableObject_OnExpose()
        {
            var host = new HostFunctions();
            var bag = new PropertyBag { { "host", host } };
            engine.AddHostObject("bag", bag);
            Assert.AreSame(engine, host.GetEngine());
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_IScriptableObject_OnModify()
        {
            var bag = new PropertyBag();
            engine.AddHostObject("bag", bag);
            var host = new HostFunctions();
            bag.Add("host", host);
            Assert.AreSame(engine, host.GetEngine());
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_Expando()
        {
            var bag = new PropertyBag();
            engine.AddHostObject("bag", bag);
            engine.Execute("bag.foo = 123");
            Assert.AreEqual(123, bag["foo"]);
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_Writable()
        {
            var bag = new PropertyBag { { "foo", false } };
            engine.AddHostObject("bag", bag);
            engine.Execute("bag.foo = true");
            Assert.IsTrue((bool)bag["foo"]);
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_Writable_Delete()
        {
            var bag = new PropertyBag { { "foo", false } };
            engine.AddHostObject("bag", bag);
            engine.Execute("delete bag.foo");
            Assert.IsFalse(bag.ContainsKey("foo"));
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_ReadOnly()
        {
            var bag = new PropertyBag(true);
            bag.SetPropertyNoCheck("foo", false);
            engine.AddHostObject("bag", bag);
            TestUtil.AssertException<UnauthorizedAccessException>(() => engine.Execute("bag.foo = true"));
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_ReadOnly_Delete()
        {
            var bag = new PropertyBag(true);
            bag.SetPropertyNoCheck("foo", false);
            engine.AddHostObject("bag", bag);
            TestUtil.AssertException<UnauthorizedAccessException>(() => engine.Execute("delete bag.foo"));
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_ExternalModification_AddProperty()
        {
            var bag = new PropertyBag { { "foo", 123 } };
            engine.AddHostObject("bag", bag);
            Assert.AreEqual(123, engine.Evaluate("bag.foo"));
            bag.Add("bar", 456);
            Assert.AreEqual(456, engine.Evaluate("bag.bar"));
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_ExternalModification_ChangeProperty()
        {
            var bag = new PropertyBag { { "foo", 123 } };
            engine.AddHostObject("bag", bag);
            Assert.AreEqual(123, engine.Evaluate("bag.foo"));
            bag["foo"] = 456;
            Assert.AreEqual(456, engine.Evaluate("bag.foo"));
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_ExternalModification_DeleteProperty()
        {
            var bag = new PropertyBag { { "foo", 123 }, { "bar", 456 } };
            engine.AddHostObject("bag", bag);
            Assert.AreEqual(456, engine.Evaluate("bag.bar"));
            bag.Remove("bar");
            Assert.AreSame(Undefined.Value, engine.Evaluate("bag.bar"));
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_MultiEngine()
        {
            var bag = new PropertyBag();
            engine.AddHostObject("bag", bag);

            Action innerTest = () =>
            {
                // The Visual Studio 2013 debugging stack fails to release the engine properly,
                // resulting in test failure. Visual Studio 2012 does not have this bug.

                using (var scriptEngine = new VBScriptEngine())
                {
                    scriptEngine.AddHostObject("bag", bag);
                    Assert.AreEqual(2, bag.EngineCount);
                }
            };

            innerTest();
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
            Assert.AreEqual(1, bag.EngineCount);
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_MultiEngine_Property()
        {
            var outerBag = new PropertyBag();
            engine.AddHostObject("bag", outerBag);

            var innerBag = new PropertyBag();
            Action innerTest = () =>
            {
                // The Visual Studio 2013 debugging stack fails to release the engine properly,
                // resulting in test failure. Visual Studio 2012 does not have this bug.

                using (var scriptEngine = new VBScriptEngine())
                {
                    scriptEngine.AddHostObject("bag", outerBag);
                    outerBag.Add("innerBag", innerBag);
                    Assert.AreEqual(2, innerBag.EngineCount);
                }
            };

            innerTest();
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
            Assert.AreEqual(1, innerBag.EngineCount);
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_MultiEngine_HostFunctions()
        {
            using (var scriptEngine = new V8ScriptEngine(V8ScriptEngineFlags.EnableDebugging))
            {
                const string code = "bag.host.func(0, function () { return bag.func(); })";
                var bag = new PropertyBag
                {
                    { "host", new HostFunctions() },
                    { "func", new Func<object>(() => ScriptEngine.Current) }
                };

                engine.AddHostObject("bag", bag);
                scriptEngine.AddHostObject("bag", bag);

                var func = (Func<object>)engine.Evaluate(code);
                Assert.AreSame(engine, func());

                func = (Func<object>)scriptEngine.Evaluate(code);
                Assert.AreSame(scriptEngine, func());
            }
        }

        [TestMethod, TestCategory("PropertyBag")]
        public void PropertyBag_MultiEngine_Parallel()
        {
            // This is a torture test for ConcurrentWeakSet and general engine teardown/cleanup.
            // It has exposed some very tricky engine bugs.

            var bag = new PropertyBag();
            engine.AddHostObject("bag", bag);

            const int threadCount = 256;
            var engineCount = 0;

            // 32-bit V8 starts failing requests to create new contexts rather quickly. This is
            // because each V8 isolate requires (among other things) a 32MB address space
            // reservation. 64-bit V8 reserves much larger blocks but benefits from the enormous
            // available address space.

            var maxV8Count = Environment.Is64BitProcess ? 128 : 16;
            var maxJScriptCount = (threadCount - maxV8Count) / 2;

            var startEvent = new ManualResetEventSlim(false);
            var checkpointEvent = new ManualResetEventSlim(false);
            var continueEvent = new ManualResetEventSlim(false);
            var stopEvent = new ManualResetEventSlim(false);

            ParameterizedThreadStart body = arg =>
            {
                // ReSharper disable AccessToDisposedClosure

                var index = (int)arg;
                startEvent.Wait();

                ScriptEngine scriptEngine;
                if (index < maxV8Count)
                {
                    scriptEngine = new V8ScriptEngine();
                }
                else if (index < (maxV8Count + maxJScriptCount))
                {
                    scriptEngine = new JScriptEngine();
                }
                else
                {
                    scriptEngine = new VBScriptEngine();
                }

                scriptEngine.AddHostObject("bag", bag);
                if (Interlocked.Increment(ref engineCount) == threadCount)
                {
                    checkpointEvent.Set();
                }

                continueEvent.Wait();

                scriptEngine.Dispose();
                if (Interlocked.Decrement(ref engineCount) == 0)
                {
                    stopEvent.Set();
                }

                // ReSharper restore AccessToDisposedClosure
            };

            var threads = Enumerable.Range(0, threadCount).Select(index => new Thread(body)).ToArray();
            threads.ForEach((thread, index) => thread.Start(index));

            startEvent.Set();
            checkpointEvent.Wait();
            Assert.AreEqual(threadCount + 1, bag.EngineCount);

            continueEvent.Set();
            stopEvent.Wait();
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
            Assert.AreEqual(1, bag.EngineCount);

            Array.ForEach(threads, thread => thread.Join());
            startEvent.Dispose();
            checkpointEvent.Dispose();
            continueEvent.Dispose();
            stopEvent.Dispose();
        }

        // ReSharper restore InconsistentNaming

        #endregion
    }
}
