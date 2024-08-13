Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports DistributionCR.Program

Namespace Tests
    <TestClass>
    Public Class TestsDistribution
        <TestMethod>
        Sub TestNumCol()
            Assert.AreEqual(0, NumCol("A"))
            Assert.AreEqual(25, NumCol("Z"))
            Assert.AreEqual(26, NumCol("AA"))
            Assert.AreEqual(27, NumCol("AB"))
        End Sub
    End Class
End Namespace
