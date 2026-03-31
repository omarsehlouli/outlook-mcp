/**
 * List emails functionality
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');

/**
 * List emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = args.count || 10;
  const toFilter = args.to || '';

  try {
    // Get access token
    const { accessToken } = await ensureAuthenticated(args.account);

    // Resolve the folder path
    const endpoint = await resolveFolderPath(accessToken, folder);

    // Add query parameters
    const queryParams = {
      $top: Math.min(50, requestedCount),
      $select: config.EMAIL_SELECT_FIELDS
    };

    // If filtering by recipient, use KQL search (can't use $orderby with $search)
    if (toFilter) {
      queryParams.$search = `"to:${toFilter}"`;
    } else {
      queryParams.$orderby = 'receivedDateTime desc';
    }

    // Make API call with pagination support
    let response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, queryParams, requestedCount);

    // If KQL search returned nothing or failed, fall back to client-side filtering
    if (toFilter && (!response.value || response.value.length === 0)) {
      const fallbackParams = {
        $top: 50,
        $orderby: 'receivedDateTime desc',
        $select: config.EMAIL_SELECT_FIELDS
      };
      // Fetch more than needed so we can filter client-side
      const allResponse = await callGraphAPIPaginated(accessToken, 'GET', endpoint, fallbackParams, Math.max(requestedCount * 5, 100));
      if (allResponse.value) {
        const filterLower = toFilter.toLowerCase();
        allResponse.value = allResponse.value.filter(email => {
          const recipients = (email.toRecipients || []).concat(email.ccRecipients || []);
          return recipients.some(r => r.emailAddress.address.toLowerCase() === filterLower);
        }).slice(0, requestedCount);
      }
      response = allResponse;
    }
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: `No emails found in ${folder}.`
        }]
      };
    }
    
    // Format results
    const emailList = response.value.map((email, index) => {
      const sender = email.from ? email.from.emailAddress : { name: 'Unknown', address: 'unknown' };
      const date = new Date(email.receivedDateTime).toLocaleString();
      const readStatus = email.isRead ? '' : '[UNREAD] ';
      
      const toAddresses = email.toRecipients
        ? email.toRecipients.map(r => r.emailAddress.address).join(', ')
        : '';

      return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\n   To: ${toAddresses}\n   Subject: ${email.subject}\n   ID: ${email.id}\n`;
    }).join("\n");
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} emails in ${folder}:\n\n${emailList}`
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
        text: `Error listing emails: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEmails;
