#!/bin/sh
#
# Licensed to the Apache Software Foundation (ASF) under one
# or more contributor license agreements.  See the NOTICE file
# distributed with this work for additional information
# regarding copyright ownership.  The ASF licenses this file
# to you under the Apache License, Version 2.0 (the
# "License"); you may not use this file except in compliance
# with the License.  You may obtain a copy of the License at
#
#   https://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied.  See the License for the
# specific language governing permissions and limitations
# under the License.

# Apache Maven Wrapper startup script, version 3.3.2

set -euf
[ "${MVNW_VERBOSE-}" != true ] || set -x

# OS specific support.
native_path() { printf '%s\n' "$1"; }
case "$(uname)" in
CYGWIN* | MINGW*)
  [ -z "${JAVA_HOME-}" ] || JAVA_HOME="$(cygpath --unix "$JAVA_HOME")"
  native_path() { cygpath --path --windows "$1"; }
  ;;
esac

# Set JAVACMD
set_java_home() {
  if [ -n "${JAVA_HOME-}" ]; then
    if [ -x "$JAVA_HOME/jre/sh/java" ]; then
      JAVACMD="$JAVA_HOME/jre/sh/java"
    else
      JAVACMD="$JAVA_HOME/bin/java"
    fi
    if [ ! -x "$JAVACMD" ] && [ -x "$JAVA_HOME/bin/java.exe" ]; then
      JAVACMD="$JAVA_HOME/bin/java.exe"
    fi
  else
    JAVACMD="$(
      'set' +e
      'unset' IFS
      command -v java
    )" || :
  fi
  if [ ! -x "${JAVACMD-}" ]; then
    echo "ERROR: JAVA_HOME not set, and no 'java' command on PATH" >&2
    exit 1
  fi
}

# Download Maven wrapper jar if missing
download_maven_wrapper() {
  wrapperJarPath=".mvn/wrapper/maven-wrapper.jar"
  if [ -r "$wrapperJarPath" ]; then
    return
  fi
  wrapperUrl=$(grep '^wrapperUrl=' ".mvn/wrapper/maven-wrapper.properties" | cut -d= -f2-)
  if [ -z "$wrapperUrl" ]; then
    echo "ERROR: wrapperUrl not set in maven-wrapper.properties" >&2
    exit 1
  fi
  echo "Downloading $wrapperUrl"
  if command -v wget > /dev/null; then
    wget -q -O "$wrapperJarPath" "$wrapperUrl" || rm -f "$wrapperJarPath"
  elif command -v curl > /dev/null; then
    curl -fsSL -o "$wrapperJarPath" "$wrapperUrl" || rm -f "$wrapperJarPath"
  else
    echo "ERROR: wget or curl required to download Maven wrapper" >&2
    exit 1
  fi
}

# Locate maven home from distribution
locate_maven_home() {
  localRepo="${MAVEN_USER_HOME:-${HOME}/.m2}"
  distUrl=$(grep '^distributionUrl=' ".mvn/wrapper/maven-wrapper.properties" | cut -d= -f2-)
  distName="${distUrl##*/}"
  distName="${distName%.zip}"
  distDir="$localRepo/wrapper/dists/$distName"
  if [ -d "$distDir" ]; then
    MAVEN_HOME=$(find "$distDir" -maxdepth 2 -name "bin" -type d 2>/dev/null | head -1 | sed 's|/bin$||')
  fi
}

# Resolve project base dir
MAVEN_PROJECTBASEDIR="${MAVEN_BASEDIR:-}"
if [ -z "${MAVEN_PROJECTBASEDIR-}" ]; then
  MAVEN_PROJECTBASEDIR="$(pwd)"
fi

set_java_home
download_maven_wrapper
locate_maven_home

if [ -n "${MAVEN_HOME-}" ] && [ -x "$MAVEN_HOME/bin/mvn" ]; then
  exec "$MAVEN_HOME/bin/mvn" "$@"
else
  exec "$JAVACMD" \
    ${JAVA_OPTS:+-Xms256m -Xmx512m} \
    -classpath ".mvn/wrapper/maven-wrapper.jar" \
    "-Dmaven.multiModuleProjectDirectory=$MAVEN_PROJECTBASEDIR" \
    org.apache.maven.wrapper.MavenWrapperMain "$@"
fi
