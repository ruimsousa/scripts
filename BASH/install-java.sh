#!/bin/bash
#================================================================
# HEADER
#================================================================
#% SYNOPSIS
#+    ${SCRIPT_NAME} arg ...
#%
#% DESCRIPTION
#% This script is used to uncompress the JAVA JDK file and
#% install it in a specific Weblogic Server environment.
#% May need to be adapted to be applied in other environments!
#%
#% OPTIONS
#%    No options implemented yet!
#%
#% EXAMPLES
#%    ${SCRIPT_NAME} arg1
#%
#================================================================
#- IMPLEMENTATION
#-    version         ${SCRIPT_NAME} ${SCRIPT_VERSION}
#-    author          eProseed Suppport Team
#-    copyright       Copyright (c) http://www.eproseed.com
#-
#================================================================
# Script: install-jdk7.sh
# Author: Support Team
#
# This script is used to uncompress the JAVA JDK file and install it
# in a specific Weblogic Server environment.
# May need to be adapted to be applied in other environments!
#================================================================
# HISTORY
#   2017/02/02 : Rui Sousa : Script creation
#   2017/02/06 : Rui Sousa : Adapting the script to expect the JDK
#                           file as a argument
#
#================================================================
# END_OF_HEADER
#================================================================

# variables
SCRIPT_VERSION=1.0.1
SCRIPT_NAME="$(basename ${0})" # scriptname without path
SCRIPT_DIR="$( cd $(dirname "$0") && pwd )" # script directory
ORACLE_USER=oracle
JAVA_HOME=/u01/oracle/products/java

# Print script information
echo "Starting ..."
echo
echo "Script Name: $SCRIPT_NAME"
echo "Script version: $VERSION"
echo "JAVA_HOME=$JAVA_HOME"
echo "ORACLE_USER=$ORACLE_USER"
echo

# Abort if not oracle user
if [[ ! `whoami` = $ORACLE_USER ]]; then
    echo "ERROR (1): You must be logged in with user oracle to run this script."
    echo "Please login with the oracle user for this environment."
    exit 1
fi

# Abort if no argument exists
# the argument should be the JDK filename with .tar.gz
if [[ -z "$1" ]]; then
    echo "ERROR (2): No arguments supplied"
    echo "Usage: $SCRIPT_NAME <path>/jdk-XuYY.tar.gz "
    exit 2
fi

# Check that the file is a JDK archive
if [[ ! $1 =~ jdk-[0-9]{1}u[0-9]{1,2}.*\.tar\.gz ]]; then
    echo "ERROR (3): '$1' doesn't look like a JDK archive."
    echo "The file name should begin 'jdk-XuYY', where X is the version number and YY is the update number."
    echo "Please re-name the file, or download the JDK again and keep the default file name."
    exit 3
else
    # check if the file exists
    if [[ ! -f $1 ]]; then
      echo "ERROR (4): Cannot find file '$1'"
      echo "Please check the path to the file and execute the script again."
      exit 4
    fi
fi

# Unpack the JDK, and get the name of the unpacked folder
echo "Uncompressing JDK file $1 to the following location: $JAVA_HOME"
tar -xf $1 -C $JAVA_HOME
JDK_VER=`echo $1 | sed -r 's/jdk-([0-9]{1}u[0-9]{1,3}).*\.tar\.gz/\1/g'`   # Version 7u131
JDK_NAME=`echo $JDK_VER | sed -r 's/([0-9]{1})u([0-9]{1,2})/jdk1.\1.0_\2/g'` # jdk1.7.0_131
echo

# remove directory symbolic link, if exists
if [[ -L "$JAVA_HOME/jdk1.7.0" ]]; then
  rm "$JAVA_HOME/jdk1.7.0"
  echo "Symbolic link $JAVA_HOME/jdk1.7.0 has been removed."
fi

# create new symbolic link
ln -s "$JAVA_HOME/$JDK_NAME" "$JAVA_HOME/jdk1.7.0"
echo "Symbolic link $JAVA_HOME/jdk1.7.0 has been created to the following target: $JAVA_HOME/$JDK_NAME."

# Copy certificates, if folder exists
if [[ -d "$JAVA_HOME/jdk1.7.0_111/jre/lib/security/cacerts" ]]; then
  echo "cacerts copied into folder $JAVA_HOME/$JDK_NAME/jre/lib/security/cacerts."
  cp -p "$JAVA_HOME/jdk1.7.0_111/jre/lib/security/cacerts" "$JAVA_HOME/$JDK_NAME/jre/lib/security/cacerts"
fi
echo
echo "Script finished."
