# Add project specific ProGuard rules here.

# Keep all data/model classes used by Gson for JSON deserialization
-keep class com.osakidetza.ope.data.** { *; }
-keep class com.osakidetza.ope.data.model.** { *; }

# Gson generic type handling
-keepattributes Signature
-keepattributes *Annotation*

# Prevent R8 from stripping Gson TypeToken and related classes
-keep class com.google.gson.** { *; }
-keep class * extends com.google.gson.TypeAdapter
-keep class * implements com.google.gson.TypeAdapterFactory
-keep class * implements com.google.gson.JsonSerializer
-keep class * implements com.google.gson.JsonDeserializer
