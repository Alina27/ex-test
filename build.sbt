name := "hw"

version := "0.1"

scalaVersion := "2.13.6"

libraryDependencies ++= Seq(
  "org.apache.poi" % "poi" % "3.9",
  "org.apache.poi" % "poi-ooxml" % "3.9"
)

scalacOptions ++= Seq("-unchecked", "-deprecation")