using System.Collections.Generic;

namespace com.microsoft.dx.officewopi.Utils
{
    /// <summary>
    /// Contains all valid URL placeholders for different WOPI actions
    /// Used to build correct action URL for iframe to WOPI
    /// </summary>
    public class WopiUrlPlaceholders
    {
        public static List<string> Placeholders = new List<string>() {
            BUSINESS_USER, DC_LLCC, DISABLE_CHAT, PERFSTATS,UI_LLCC,
            HOST_SESSION_ID, SESSION_CONTEXT, WOPI_SOURCE, ACTIVITY_NAVIGATION_ID,
            DISABLE_ASYNC, DISABLE_BROADCAST, EMBDDED, FULLSCREEN,
            RECORDING, THEME_ID, VALIDATOR_TEST_CATEGORY
        };

        public const string UI_LLCC = "<ui=UI_LLCC&>";
        public const string DC_LLCC = "<rs=DC_LLCC&>";
        public const string DISABLE_CHAT = "<dchat=DISABLE_CHAT&>";
        public const string HOST_SESSION_ID = "<hid=HOST_SESSION_ID&>";
        public const string SESSION_CONTEXT = "<sc=SESSION_CONTEXT&>";
        public const string WOPI_SOURCE = "<wopisrc=WOPI_SOURCE&>";
        public const string PERFSTATS = "<showpagestats=PERFSTATS&>";
        public const string BUSINESS_USER = "<IsLicensedUser=BUSINESS_USER&>";
        public const string ACTIVITY_NAVIGATION_ID = "<actnavid=ACTIVITY_NAVIGATION_ID&>";

        public const string DISABLE_ASYNC = "<na=DISABLE_ASYNC&>";
        public const string DISABLE_BROADCAST = "<vp=DISABLE_BROADCAST&>";
        public const string EMBDDED = "<e=EMBEDDED&>";
        public const string FULLSCREEN = "<fs=FULLSCREEN&>";
        public const string RECORDING = "<rec=RECORDING&>";
        public const string THEME_ID = "<thm=THEME_ID&>";
        public const string VALIDATOR_TEST_CATEGORY = "<testcategory=VALIDATOR_TEST_CATEGORY>";
    }



    /// <summary>
    /// Constains paths and confiurations to handle WOPI requests
    /// </summary>
    public static class WopiRequestConst
    {
        //WOPI protocol constants
        public const string WOPI_BASE_PATH = @"/wopi/";
        public const string WOPI_CHILDREN_PATH = @"/children";
        public const string WOPI_CONTENTS_PATH = @"/contents";
        public const string WOPI_FILES_PATH = @"files/";
        public const string WOPI_FOLDERS_PATH = @"folders/";
    }

    /// <summary>
    /// Contains valid WOPI request headers
    /// </summary>
    public class WopiRequestHeaders
    {
        //WOPI Header Consts
        public const string APP_ENDPOINT = "X-WOPI-AppEndpoint";
        public const string CLIENT_VERSION = "X-WOPI-ClientVersion";
        public const string CORRELATION_ID = "X-WOPI-CorrelationId";
        public const string LOCK = "X-WOPI-Lock";
        public const string MACHINE_NAME = "X-WOPI-MachineName";
        public const string MAX_EXPECTED_SIZE = "X-WOPI-MaxExpectedSize";
        public const string OLD_LOCK = "X-WOPI-OldLock";
        public const string OVERRIDE = "X-WOPI-Override";
        public const string OVERWRITE_RELATIVE_TARGET = "X-WOPI-OverwriteRelativeTarget";
        public const string PREF_TRACE_REQUESTED = "X-WOPI-PerfTraceRequested";
        public const string PROOF = "X-WOPI-Proof";
        public const string PROOF_OLD = "X-WOPI-ProofOld";
        public const string RELATIVE_TARGET = "X-WOPI-RelativeTarget";
        public const string REQUESTED_NAME = "X-WOPI-RequestedName";
        public const string SESSION_CONTEXT = "X-WOPI-SessionContext";
        public const string SIZE = "X-WOPI-Size";
        public const string SUGGESTED_TARGET = "X-WOPI-SuggestedTarget";
        public const string TIME_STAMP = "X-WOPI-TimeStamp";
    }

    /// <summary>
    /// Contains valid WOPI response headers
    /// </summary>
    public class WopiResponseHeaders
    {
        //WOPI Header Consts
        public const string HOST_ENDPOINT = "X-WOPI-HostEndpoint";
        public const string INVALID_FILE_NAME_ERROR = "X-WOPI-InvalidFileNameError";
        public const string LOCK = "X-WOPI-Lock";
        public const string LOCK_FAILURE_REASON = "X-WOPI-LockFailureReason";
        public const string LOCKED_BY_OTHER_INTERFACE = "X-WOPI-LockedByOtherInterface";
        public const string MACHINE_NAME = "X-WOPI-MachineName";
        public const string PREF_TRACE = "X-WOPI-PerfTrace";
        public const string SERVER_ERROR = "X-WOPI-ServerError";
        public const string SERVER_VERSION = "X-WOPI-ServerVersion";
        public const string VALID_RELATIVE_TARGET = "X-WOPI-ValidRelativeTarget";
    }
}