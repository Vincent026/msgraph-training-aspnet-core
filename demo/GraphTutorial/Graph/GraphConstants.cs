// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <GraphConstantsSnippet>
namespace GraphTutorial
{
    public static class GraphConstants
    {
        // Defines the permission scopes used by the app
        public readonly static string[] Scopes =
        {
            "User.Read",
            "MailboxSettings.Read",
            "Mail.ReadBasic",
            "Files.Read",
            "Calendars.ReadWrite"
        };
    }
}
// </GraphConstantsSnippet>
