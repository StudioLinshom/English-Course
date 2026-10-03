#!/bin/sh
# Builds english-words.apk without Gradle, using Ubuntu's packaged tools:
#   apt-get install aapt apksigner dalvik-exchange zipalign android-sdk-platform-23
# Usage: KEYSTORE=/path/release.jks KEYPASS=... ./build.sh
set -e
cd "$(dirname "$0")"
: "${KEYSTORE:?set KEYSTORE to the signing keystore}" "${KEYPASS:?set KEYPASS}"
SDK=/usr/lib/android-sdk
JAR=$SDK/platforms/android-23/android.jar
OUT=build
rm -rf $OUT && mkdir -p $OUT/gen $OUT/classes

aapt package -f -m -J $OUT/gen -M AndroidManifest.xml -S res -I $JAR \
  --min-sdk-version 23 --target-sdk-version 34 -F $OUT/unsigned.apk

javac -source 8 -target 8 -nowarn -Xlint:-options -bootclasspath $JAR \
  -d $OUT/classes $(find src $OUT/gen -name '*.java')

dalvik-exchange --dex --min-sdk-version=23 --output=$OUT/classes.dex $OUT/classes
(cd $OUT && aapt add unsigned.apk classes.dex >/dev/null)

zipalign -f -p 4 $OUT/unsigned.apk $OUT/aligned.apk
apksigner sign --ks "$KEYSTORE" --ks-pass env:KEYPASS --key-pass env:KEYPASS --v4-signing-enabled false \
  --out ../english-words.apk $OUT/aligned.apk
apksigner verify ../english-words.apk
echo "Built $(cd .. && pwd)/english-words.apk"
