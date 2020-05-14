import jaydebeapi
import jpype
import os,sys
import ctypes

os.environ["JAVA_HOME"] = "C:/Users/z659190/Documents/Run/jdk/"
print ("setenv JAVA_HOME", os.environ["JAVA_HOME"])

jvmPath=r'C:/Users/z659190/Documents/Run/jdk/bin/server/jvm.dll'

# print(jvmPath = jpype.getDefaultJVMPath() )
jpype.startJVM(jvmpath=jvmPath,convertStrings=True)
# jpype.startJVM(jvmpath=jvmpath,classpath='C:/Program Files (x86)/Java/jre1.8.0_152/lib/')
jpype.java.lang.System.out.println("hello world!")

# ctypes.CDLL(r'C:\Program Files (x86)\Java\jre1.8.0_152\bin\client\jvm.dll')
driver = "de.simplicit.vjdbc.VirtualDriverBirtWrapper"
# url = "jdbc:vjdbc:servlet:https://zffriedric-stage.plateau.com/vjdbc/vjdbc,db10g"
url = "jdbc:vjdbc:servlet@https://zffriedric-stage.plateau.com/vjdbc/vjdbc:db10g"
jarfile = os.path.abspath('.')+"/birtvjdbcdriver.jar"
sys.path.append(jarfile)
# "c:/Users/z659190/Documents/Run/Plateau Report Designer/plugins/org.eclipse.birt.report.data.oda.jdbc_4.4.1.v201409050910/drivers/birtvjdbcdriver.jar",)

conn = jaydebeapi.connect(driver,[url,"10104100", "Password1"],jarfile)
curs = conn.cursor()
#
curs.execute("select * from PA_STUDENT")
curs.fetchall()
curs.close()
conn.close()


jpype.shutdownJVM()