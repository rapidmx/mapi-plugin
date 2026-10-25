// Re-exports just the library's MongoDB model classes MAPI needs so the test Server's ClassLoader
// (rooted at `test/server-mongo`) can discover their `@DataStore` metadata alongside the test routes that
// use them. Deliberately a NAMED (not wildcard) re-export: `@rapidmx/restapi/mongo` bundles routes/jobs
// alongside its models, and a wildcard re-export here would make the ClassLoader also discover and eagerly
// initialize every REST route/job class restapi defines - none of which this MAPI-only test harness
// configures dependencies for.
export {
    AuditLogEntryMongo,
    CalendarEventMongo,
    EncryptionPolicyMongo,
    ContactMongo,
    FolderMongo,
    LabelMongo,
    MailboxMongo,
    MessageMongo,
    TaskMongo,
    ZeroAccessIndexEntryMongo,
} from "@rapidmx/restapi/mongo";
