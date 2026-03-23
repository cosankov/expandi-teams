// Microsoft Teams Adaptive Card formatter for Expandi webhook events
// Docs: https://learn.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/connectors-using

const EVENT_LABELS = {
  'linked_in_messenger.campaign_replied': { emoji: '\u{1F4AC}', label: 'New Reply', color: 'accent' },
  'linked_in_messenger.campaign_connected': { emoji: '\u{1F91D}', label: 'Connection Accepted', color: 'good' },
  'linked_in_messenger.campaign_message_sent': { emoji: '\u{1F4E4}', label: 'Message Sent', color: 'default' },
  'linked_in_messenger.campaign_invite_sent': { emoji: '\u{1F4E8}', label: 'Connection Request Sent', color: 'default' },
  'linked_in_messenger.campaign_profile_visited': { emoji: '\u{1F441}', label: 'Profile Visited', color: 'default' },
  'linked_in_messenger.campaign_post_liked': { emoji: '\u{1F44D}', label: 'Post Liked', color: 'default' },
  'linked_in_messenger.campaign_follow_sent': { emoji: '\u{2795}', label: 'Follow Sent', color: 'default' },
  'linked_in_messenger.campaign_company_follow_sent': { emoji: '\u{1F3E2}', label: 'Company Follow Sent', color: 'default' },
  'linked_in_messenger.campaign_endorsement_sent': { emoji: '\u{2B50}', label: 'Endorsement Sent', color: 'default' },
  'linked_in_messenger.campaign_contact_tagged': { emoji: '\u{1F3F7}', label: 'Contact Tagged', color: 'default' },
  'linked_in_messenger.campaign_contact_disconnected': { emoji: '\u{1F50C}', label: 'Contact Disconnected', color: 'attention' },
  'linked_in_messenger.campaign_contact_revoked': { emoji: '\u{274C}', label: 'Connection Revoked', color: 'attention' },
  'linked_in_messenger.campaign_finished': { emoji: '\u{2705}', label: 'Campaign Finished', color: 'good' },
  'linked_in_messenger.campaign_email_sent': { emoji: '\u{2709}\u{FE0F}', label: 'Email Sent', color: 'default' },
  'linked_in_messenger.campaign_email_opened': { emoji: '\u{1F4EC}', label: 'Email Opened', color: 'default' },
  'linked_in_messenger.campaign_email_clicked': { emoji: '\u{1F517}', label: 'Email Link Clicked', color: 'default' },
  'linked_in_messenger.campaign_email_bounced': { emoji: '\u{26A0}\u{FE0F}', label: 'Email Bounced', color: 'warning' },
  'linked_in_messenger.campaign_first_reply': { emoji: '\u{1F3AF}', label: 'First Reply', color: 'accent' },
  'linked_in_messenger.no_connection_requests_scheduled': { emoji: '\u{1F4CB}', label: 'No Connection Requests Scheduled', color: 'warning' },
  'linked_in_messenger.no_messages_scheduled': { emoji: '\u{1F4CB}', label: 'No Messages Scheduled', color: 'warning' },
  'linked_in_messenger.nothing_scheduled': { emoji: '\u{1F4CB}', label: 'Nothing Scheduled', color: 'warning' },
};

function getEventInfo(event) {
  return EVENT_LABELS[event] || { emoji: '\u{1F514}', label: event.split('.').pop().replace(/_/g, ' '), color: 'default' };
}

function isValidUrl(str) {
  if (!str || typeof str !== 'string') return false;
  return str.startsWith('https://') || str.startsWith('http://');
}

function truncate(str, max = 300) {
  if (!str) return '';
  return str.length > max ? str.substring(0, max) + '...' : str;
}

function formatDate(dateStr) {
  if (!dateStr) return 'N/A';
  const d = new Date(dateStr);
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
}

// --- Adaptive Card builders ---

function makeHeaderRow(eventInfo, campaign, sender) {
  return {
    type: 'ColumnSet',
    columns: [
      {
        type: 'Column',
        width: 'stretch',
        items: [
          {
            type: 'TextBlock',
            text: `${eventInfo.emoji} **${eventInfo.label}**`,
            size: 'large',
            weight: 'bolder',
            wrap: true,
          },
          {
            type: 'TextBlock',
            text: `Campaign: ${campaign}  \u00B7  Sender: ${sender}`,
            size: 'small',
            isSubtle: true,
            spacing: 'none',
            wrap: true,
          },
        ],
      },
    ],
  };
}

function makeContactRow(name, title, companyLine, contactInfo, imageUrl) {
  const column = {
    type: 'Column',
    width: 'stretch',
    items: [],
  };

  column.items.push({
    type: 'TextBlock',
    text: `**${name}**`,
    size: 'medium',
    weight: 'bolder',
    wrap: true,
  });

  if (title) {
    column.items.push({
      type: 'TextBlock',
      text: title,
      size: 'small',
      spacing: 'none',
      wrap: true,
    });
  }

  if (companyLine) {
    column.items.push({
      type: 'TextBlock',
      text: companyLine,
      size: 'small',
      isSubtle: true,
      spacing: 'none',
      wrap: true,
    });
  }

  if (contactInfo) {
    column.items.push({
      type: 'TextBlock',
      text: contactInfo,
      size: 'small',
      isSubtle: true,
      spacing: 'none',
      wrap: true,
    });
  }

  const columns = [column];

  if (isValidUrl(imageUrl)) {
    columns.push({
      type: 'Column',
      width: 'auto',
      items: [
        {
          type: 'Image',
          url: imageUrl,
          size: 'small',
          style: 'person',
        },
      ],
    });
  }

  return { type: 'ColumnSet', columns };
}

function makeQuoteBlock(emoji, label, text) {
  return [
    {
      type: 'TextBlock',
      text: `${emoji} **${label}:**`,
      size: 'small',
      weight: 'bolder',
      spacing: 'medium',
      wrap: true,
    },
    {
      type: 'TextBlock',
      text: truncate(text),
      size: 'small',
      isSubtle: true,
      wrap: true,
      spacing: 'small',
    },
  ];
}

function makeLinks(links) {
  if (links.length === 0) return null;
  const parts = links.map(l => `[${l.label}](${l.url})`);
  return {
    type: 'TextBlock',
    text: parts.join('  \u00B7  '),
    size: 'small',
    spacing: 'medium',
    wrap: true,
  };
}

// --- Event-specific formatters ---

function formatReply(payload) {
  const { hook, contact, messenger, campaign_instance_contact } = payload;
  const eventInfo = getEventInfo(hook.event);
  const campaign = messenger?.campaign_instance || 'Unknown Campaign';
  const sender = hook.li_account_name || 'Unknown';

  const name = [contact?.first_name, contact?.last_name].filter(Boolean).join(' ') || 'Unknown';
  const title = contact?.job_title || '';
  const company = contact?.company;

  let companyLine = '';
  if (company) {
    const parts = [company.name];
    if (company.employee_count_start && company.employee_count_end) {
      parts.push(`${company.employee_count_start}-${company.employee_count_end} employees`);
    }
    if (company.location) parts.push(company.location);
    companyLine = parts.join(' \u00B7 ');
  } else if (contact?.company_name) {
    companyLine = contact.company_name;
  }

  const body = [];

  // Header
  body.push(makeHeaderRow(eventInfo, campaign, sender));

  // Separator
  body.push({ type: 'TextBlock', text: ' ', size: 'small', spacing: 'small' });

  // Contact info (email/phone)
  const contactParts = [];
  if (contact?.email) contactParts.push(`\u{2709}\u{FE0F} ${contact.email}`);
  if (contact?.phone) contactParts.push(`\u{1F4DE} ${contact.phone}`);
  const contactInfo = contactParts.length > 0 ? contactParts.join('  \u00B7  ') : '';

  // Contact
  body.push(makeContactRow(name, title, companyLine, contactInfo, contact?.image_link));

  // Their reply
  if (messenger?.last_received_message) {
    body.push(...makeQuoteBlock('\u{1F4AC}', 'Their reply', messenger.last_received_message));
  }

  // What you sent
  if (messenger?.last_sent_message) {
    body.push(...makeQuoteBlock('\u{1F4E4}', 'You sent', messenger.last_sent_message));
  }

  // Stats
  const stats = [];
  if (campaign_instance_contact?.nr_steps_before_responding) {
    stats.push(`Steps before reply: **${campaign_instance_contact.nr_steps_before_responding}**`);
  }
  if (messenger?.connected_at) {
    stats.push(`Connected: **${formatDate(messenger.connected_at)}**`);
  }
  if (messenger?.last_received_message_datetime) {
    stats.push(`Replied: **${formatDate(messenger.last_received_message_datetime)}**`);
  }
  if (stats.length > 0) {
    body.push({
      type: 'TextBlock',
      text: stats.join('  \u00B7  '),
      size: 'small',
      spacing: 'medium',
      wrap: true,
    });
  }

  // Links
  const links = [];
  if (isValidUrl(contact?.profile_link)) links.push({ label: '\u{1F517} LinkedIn', url: contact.profile_link });
  if (isValidUrl(messenger?.lead_inbox_link)) links.push({ label: '\u{1F4E5} Expandi Inbox', url: messenger.lead_inbox_link });
  if (isValidUrl(contact?.sales_nav_link)) links.push({ label: '\u{1F50E} Sales Nav', url: contact.sales_nav_link });
  if (isValidUrl(messenger?.thread)) links.push({ label: '\u{1F4AC} Thread', url: messenger.thread });
  const linksBlock = makeLinks(links);
  if (linksBlock) body.push(linksBlock);

  return body;
}

function formatConnection(payload) {
  const { hook, contact, messenger } = payload;
  const eventInfo = getEventInfo(hook.event);
  const campaign = messenger?.campaign_instance || 'Unknown Campaign';
  const sender = hook.li_account_name || 'Unknown';

  const name = [contact?.first_name, contact?.last_name].filter(Boolean).join(' ') || 'Unknown';
  const title = contact?.job_title || '';
  const company = contact?.company;

  let companyLine = '';
  if (company) {
    const parts = [company.name];
    if (company.employee_count_start && company.employee_count_end) {
      parts.push(`${company.employee_count_start}-${company.employee_count_end} employees`);
    }
    if (company.location) parts.push(company.location);
    companyLine = parts.join(' \u00B7 ');
  } else if (contact?.company_name) {
    companyLine = contact.company_name;
  }

  const contactParts = [];
  if (contact?.email) contactParts.push(`\u{2709}\u{FE0F} ${contact.email}`);
  if (contact?.phone) contactParts.push(`\u{1F4DE} ${contact.phone}`);
  const contactInfo = contactParts.length > 0 ? contactParts.join('  \u00B7  ') : '';

  const body = [];
  body.push(makeHeaderRow(eventInfo, campaign, sender));
  body.push({ type: 'TextBlock', text: ' ', size: 'small', spacing: 'small' });
  body.push(makeContactRow(name, title, companyLine, contactInfo, contact?.image_link));

  if (messenger?.connected_at) {
    body.push({
      type: 'TextBlock',
      text: `Connected: **${formatDate(messenger.connected_at)}**`,
      size: 'small',
      spacing: 'medium',
      wrap: true,
    });
  }

  const links = [];
  if (isValidUrl(contact?.profile_link)) links.push({ label: '\u{1F517} LinkedIn', url: contact.profile_link });
  if (isValidUrl(messenger?.lead_inbox_link)) links.push({ label: '\u{1F4E5} Expandi Inbox', url: messenger.lead_inbox_link });
  if (isValidUrl(contact?.sales_nav_link)) links.push({ label: '\u{1F50E} Sales Nav', url: contact.sales_nav_link });
  const linksBlock = makeLinks(links);
  if (linksBlock) body.push(linksBlock);

  return body;
}

function formatGenericAction(payload) {
  const { hook, contact, messenger } = payload;
  const eventInfo = getEventInfo(hook.event);
  const campaign = messenger?.campaign_instance || 'Unknown Campaign';
  const sender = hook.li_account_name || 'Unknown';

  const name = [contact?.first_name, contact?.last_name].filter(Boolean).join(' ') || 'Unknown';

  const body = [];
  body.push(makeHeaderRow(eventInfo, campaign, sender));
  body.push({ type: 'TextBlock', text: ' ', size: 'small', spacing: 'small' });

  let companyLine = '';
  if (contact?.company?.name) companyLine = contact.company.name;

  const contactParts = [];
  if (contact?.email) contactParts.push(`\u{2709}\u{FE0F} ${contact.email}`);
  if (contact?.phone) contactParts.push(`\u{1F4DE} ${contact.phone}`);
  const contactInfo = contactParts.length > 0 ? contactParts.join('  \u00B7  ') : '';

  body.push(makeContactRow(name, contact?.job_title || '', companyLine, contactInfo, contact?.image_link));

  const links = [];
  if (isValidUrl(contact?.profile_link)) links.push({ label: '\u{1F517} LinkedIn', url: contact.profile_link });
  if (isValidUrl(messenger?.lead_inbox_link)) links.push({ label: '\u{1F4E5} Expandi Inbox', url: messenger.lead_inbox_link });
  const linksBlock = makeLinks(links);
  if (linksBlock) body.push(linksBlock);

  return body;
}

function formatSystemEvent(payload) {
  const { hook } = payload;
  const eventInfo = getEventInfo(hook.event);
  const sender = hook.li_account_name || 'Unknown';

  return [
    makeHeaderRow(eventInfo, '', sender),
    {
      type: 'TextBlock',
      text: `Account: **${sender}**`,
      size: 'small',
      spacing: 'medium',
      wrap: true,
    },
  ];
}

// --- Main formatter ---

function formatExpandiPayload(payload) {
  const event = payload?.hook?.event || '';

  if (event.includes('replied') || event.includes('first_reply')) {
    return formatReply(payload);
  }

  if (event.includes('connected') || event.includes('invite_sent') || event.includes('revoked') || event.includes('disconnected')) {
    return formatConnection(payload);
  }

  if (event.includes('no_') || event.includes('nothing_scheduled')) {
    return formatSystemEvent(payload);
  }

  return formatGenericAction(payload);
}

function buildTeamsPayload(payload) {
  const body = formatExpandiPayload(payload);

  // Adaptive Card wrapped in the Teams webhook envelope
  return {
    type: 'message',
    attachments: [
      {
        contentType: 'application/vnd.microsoft.card.adaptive',
        contentUrl: null,
        content: {
          $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
          type: 'AdaptiveCard',
          version: '1.4',
          body,
        },
      },
    ],
  };
}

module.exports = { buildTeamsPayload, getEventInfo };
