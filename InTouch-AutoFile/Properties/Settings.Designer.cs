﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace InTouch_AutoFile.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "17.3.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string LastContactFolder {
            get {
                return ((string)(this["LastContactFolder"]));
            }
            set {
                this["LastContactFolder"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool TaskInbox {
            get {
                return ((bool)(this["TaskInbox"]));
            }
            set {
                this["TaskInbox"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool TaskSent {
            get {
                return ((bool)(this["TaskSent"]));
            }
            set {
                this["TaskSent"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool TaskDuplicates {
            get {
                return ((bool)(this["TaskDuplicates"]));
            }
            set {
                this["TaskDuplicates"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool TaskEmailRouting {
            get {
                return ((bool)(this["TaskEmailRouting"]));
            }
            set {
                this["TaskEmailRouting"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("dallasadams@hotmail.com; inta@bigpond.net.au; dallasadams@outlook.com; dallas.ada" +
            "ms@outlook.com.au; dallasadams@intacomputers.com; know_one@outlook.com.au; sloth" +
            "ie@outlook.com.au; dallas.a.adams@gmail.com;")]
        public string EmailRoutingAddresses {
            get {
                return ((string)(this["EmailRoutingAddresses"]));
            }
            set {
                this["EmailRoutingAddresses"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("2022-10-01")]
        public global::System.DateTime LastAliasCheck {
            get {
                return ((global::System.DateTime)(this["LastAliasCheck"]));
            }
            set {
                this["LastAliasCheck"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string CurrentAliasGUID {
            get {
                return ((string)(this["CurrentAliasGUID"]));
            }
            set {
                this["CurrentAliasGUID"] = value;
            }
        }
    }
}
