﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Stage3_FileAttributeWriteToDestination.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.6.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://m365x938597.sharepoint.com/sites/sourceSite")]
        public string SourceSPOSiteURL {
            get {
                return ((string)(this["SourceSPOSiteURL"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("source1")]
        public string DocLibDisplayName {
            get {
                return ((string)(this["DocLibDisplayName"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string LisLibDisplayName {
            get {
                return ((string)(this["LisLibDisplayName"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("admin@M365x938597.onmicrosoft.com")]
        public string Office365Username {
            get {
                return ((string)(this["Office365Username"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("B5TBg8Q4KX")]
        public string Office365Password {
            get {
                return ((string)(this["Office365Password"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("c:\\\\export\\\\source.csv")]
        public string MetadataFileExportLocation {
            get {
                return ((string)(this["MetadataFileExportLocation"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("100")]
        public string SPOQueryBatchSize {
            get {
                return ((string)(this["SPOQueryBatchSize"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://m365x907420.sharepoint.com/sites/destinationSite")]
        public string DestinationSPOSiteURL {
            get {
                return ((string)(this["DestinationSPOSiteURL"]));
            }
            set {
                this["DestinationSPOSiteURL"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Destination1")]
        public string DestinationDocLibDisplayName {
            get {
                return ((string)(this["DestinationDocLibDisplayName"]));
            }
            set {
                this["DestinationDocLibDisplayName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string DestinationLisLibDisplayName {
            get {
                return ((string)(this["DestinationLisLibDisplayName"]));
            }
            set {
                this["DestinationLisLibDisplayName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("admin@M365x907420.onmicrosoft.com")]
        public string DestinationOffice365Username {
            get {
                return ((string)(this["DestinationOffice365Username"]));
            }
            set {
                this["DestinationOffice365Username"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("8so49UBbHO")]
        public string DestinationOffice365Password {
            get {
                return ((string)(this["DestinationOffice365Password"]));
            }
            set {
                this["DestinationOffice365Password"] = value;
            }
        }
    }
}
