/**
 * Email rules management module for Outlook MCP server
 */
const handleListRules = require('./list');
const handleCreateRule = require('./create');

// Import getInboxRules for the edit sequence tool
const { getInboxRules } = require('./list');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Edit rule sequence handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleEditRuleSequence(args) {
  const { ruleName, sequence } = args;
  
  if (!ruleName) {
    return {
      content: [{ 
        type: "text", 
        text: "Rule name is required. Please specify the exact name of an existing rule."
      }]
    };
  }
  
  if (!sequence || isNaN(sequence) || sequence < 1) {
    return {
      content: [{ 
        type: "text", 
        text: "A positive sequence number is required. Lower numbers run first (higher priority)."
      }]
    };
  }
  
  try {
    // Get access token
    const { accessToken } = await ensureAuthenticated(args.account);
    
    // Get all rules
    const rules = await getInboxRules(accessToken);
    
    // Find the rule by name
    const rule = rules.find(r => r.displayName === ruleName);
    if (!rule) {
      return {
        content: [{ 
          type: "text", 
          text: `Rule with name "${ruleName}" not found.`
        }]
      };
    }
    
    // Update the rule sequence
    const updateResult = await callGraphAPI(
      accessToken,
      'PATCH',
      `me/mailFolders/inbox/messageRules/${rule.id}`,
      {
        sequence: sequence
      }
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `Successfully updated the sequence of rule "${ruleName}" to ${sequence}.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error updating rule sequence: ${error.message}`
      }]
    };
  }
}

// Rules management tool definitions
const rulesTools = [
  {
    name: "list-rules",
    description: "Lists inbox rules in your Outlook account",
    inputSchema: {
      type: "object",
      properties: {
        includeDetails: {
          type: "boolean",
          description: "Include detailed rule conditions and actions"
        },
        account: {
          type: "string",
          description: "Email account (primary email address). Omit for default account."
        }
      },
      required: []
    },
    handler: handleListRules
  },
  {
    name: "create-rule",
    description: "Creates a new inbox rule",
    inputSchema: {
      type: "object",
      properties: {
        name: {
          type: "string",
          description: "Name of the rule to create"
        },
        // --- Conditions: People ---
        fromAddresses: {
          type: "string",
          description: "Comma-separated list of sender email addresses for the rule"
        },
        sentToAddresses: {
          type: "string",
          description: "Comma-separated list of recipient (To) email addresses the email must be sent to"
        },
        senderContains: {
          type: "string",
          description: "Comma-separated strings to match in the sender address (e.g. domain names like 'contoso.com')"
        },
        recipientContains: {
          type: "string",
          description: "Comma-separated strings to match in To or Cc recipient addresses"
        },
        // --- Conditions: My name is ---
        sentToMe: {
          type: "boolean",
          description: "Match when I'm on the To line"
        },
        sentCcMe: {
          type: "boolean",
          description: "Match when I'm on the Cc line"
        },
        sentToOrCcMe: {
          type: "boolean",
          description: "Match when I'm on the To or Cc line"
        },
        notSentToMe: {
          type: "boolean",
          description: "Match when I'm not on the To line"
        },
        sentOnlyToMe: {
          type: "boolean",
          description: "Match when I'm the only recipient"
        },
        // --- Conditions: Subject/Keywords ---
        containsSubject: {
          type: "string",
          description: "Subject text the email must contain"
        },
        bodyContains: {
          type: "string",
          description: "Text the email body must contain"
        },
        bodyOrSubjectContains: {
          type: "string",
          description: "Text that must appear in either the subject or body of the email"
        },
        headerContains: {
          type: "string",
          description: "Comma-separated strings to match in message headers"
        },
        // --- Conditions: Marked with ---
        importance: {
          type: "string",
          enum: ["low", "normal", "high"],
          description: "Match emails with this importance level"
        },
        sensitivity: {
          type: "string",
          enum: ["normal", "personal", "private", "confidential"],
          description: "Match emails with this sensitivity level"
        },
        // --- Conditions: Message includes ---
        hasAttachments: {
          type: "boolean",
          description: "Whether the rule applies to emails with attachments"
        },
        categories: {
          type: "string",
          description: "Comma-separated category names the email must be labeled with"
        },
        messageActionFlag: {
          type: "string",
          enum: ["any", "call", "doNotForward", "followUp", "fyi", "forward", "noResponseNecessary", "read", "reply", "replyToAll", "review"],
          description: "Match emails with this flag-for-action value"
        },
        // --- Conditions: Message size ---
        withinSizeRangeMin: {
          type: "number",
          description: "Minimum message size in kilobytes"
        },
        withinSizeRangeMax: {
          type: "number",
          description: "Maximum message size in kilobytes"
        },
        // --- Actions ---
        moveToFolder: {
          type: "string",
          description: "Name of the folder to move matching emails to"
        },
        copyToFolder: {
          type: "string",
          description: "Name of the folder to copy matching emails to"
        },
        markAsRead: {
          type: "boolean",
          description: "Whether to mark matching emails as read"
        },
        markImportance: {
          type: "string",
          enum: ["low", "normal", "high"],
          description: "Set the importance level on matching emails"
        },
        forwardTo: {
          type: "string",
          description: "Comma-separated email addresses to forward matching emails to"
        },
        redirectTo: {
          type: "string",
          description: "Comma-separated email addresses to redirect matching emails to"
        },
        deleteMessage: {
          type: "boolean",
          description: "Whether to move matching emails to Deleted Items"
        },
        stopProcessingRules: {
          type: "boolean",
          description: "Whether to stop processing subsequent rules after this one matches"
        },
        assignCategories: {
          type: "string",
          description: "Comma-separated category names to assign to matching emails"
        },
        // --- Meta ---
        isEnabled: {
          type: "boolean",
          description: "Whether the rule should be enabled after creation (default: true)"
        },
        sequence: {
          type: "number",
          description: "Order in which the rule is executed (lower numbers run first, default: 100)"
        },
        account: {
          type: "string",
          description: "Email account to create the rule on (primary email address). Omit for default account."
        }
      },
      required: ["name"]
    },
    handler: handleCreateRule
  },
  {
    name: "edit-rule-sequence",
    description: "Changes the execution order of an existing inbox rule",
    inputSchema: {
      type: "object",
      properties: {
        ruleName: {
          type: "string",
          description: "Name of the rule to modify"
        },
        sequence: {
          type: "number",
          description: "New sequence value for the rule (lower numbers run first)"
        },
        account: {
          type: "string",
          description: "Email account (primary email address). Omit for default account."
        }
      },
      required: ["ruleName", "sequence"]
    },
    handler: handleEditRuleSequence
  }
];

module.exports = {
  rulesTools,
  handleListRules,
  handleCreateRule,
  handleEditRuleSequence
};
