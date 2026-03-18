import * as XLSX from 'xlsx';

export function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const cases = raw.map((row) => ({
          controlNumber: String(row['Control #'] || ''),
          frAssigned: row['FR Assigned'] || '',
          customerName: row['Customer Name #'] || '',
          surveyType: row['Survey Type'] || '',
          dateOrdered: row['Date Ordered'] || '',
          dateLastContacted: row['Date Last Contacted'] || '',
          appointmentDate: row['Appointment Date'] || '',
          dateReturned: row['Date Returned'] || '',
          customerDueDate: row['Customer Due Date'] || '',
          fieldStatus: row['Field Status'] || '',
          lastNote: row['Last Note'] || '',
          address: row['Address'] || '',
          city: row['City'] || '',
          state: row['State'] || '',
          zip: String(row['Zip'] || '').padStart(5, '0'),
        }));
        resolve(cases);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export function parseDate(str) {
  if (!str) return null;
  const cleaned = str.replace(/\s+\d{1,2}:\d{2}\s*(AM|PM)?/i, '').trim();
  const parts = cleaned.split('/');
  if (parts.length !== 3) return null;
  const m = parseInt(parts[0], 10) - 1;
  const d = parseInt(parts[1], 10);
  const y = parseInt(parts[2], 10);
  return new Date(y, m, d);
}

export function getDaysUntilDue(dueDateStr) {
  const due = parseDate(dueDateStr);
  if (!due) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  due.setHours(0, 0, 0, 0);
  return Math.ceil((due - today) / (1000 * 60 * 60 * 24));
}

export function getUrgency(dueDateStr) {
  const days = getDaysUntilDue(dueDateStr);
  if (days === null) return 'unknown';
  if (days < 0) return 'overdue';
  if (days <= 3) return 'due-soon';
  return 'on-track';
}

export function getUrgencyLabel(dueDateStr) {
  const days = getDaysUntilDue(dueDateStr);
  if (days === null) return 'No due date';
  if (days < 0) return `${Math.abs(days)} day${Math.abs(days) !== 1 ? 's' : ''} overdue`;
  if (days === 0) return 'Due today';
  if (days === 1) return 'Due tomorrow';
  return `${days} days left`;
}

export function getFrName(frAssigned) {
  const match = frAssigned.match(/^([^(]+)/);
  return match ? match[1].trim() : frAssigned;
}

export function getFrId(frAssigned) {
  const match = frAssigned.match(/\((\d+)\)/);
  return match ? match[1] : '';
}

export function groupByFR(cases) {
  const groups = {};
  cases.forEach((c) => {
    const name = getFrName(c.frAssigned);
    if (!groups[name]) groups[name] = [];
    groups[name].push(c);
  });
  return groups;
}

export function generateTextReminder(caseItem) {
  const frName = getFrName(caseItem.frAssigned).split(' ')[0];
  const urgency = getUrgency(caseItem.customerDueDate);
  const daysInfo = getUrgencyLabel(caseItem.customerDueDate);

  let urgencyLine = '';
  if (urgency === 'overdue') {
    urgencyLine = `URGENT: This case is ${daysInfo}. `;
  } else if (urgency === 'due-soon') {
    urgencyLine = `REMINDER: ${daysInfo} until the due date. `;
  }

  return `Hi ${frName}, this is a reminder regarding Control #${caseItem.controlNumber} (${caseItem.surveyType}) at ${caseItem.address}, ${caseItem.city}, ${caseItem.state} ${caseItem.zip}. ${urgencyLine}This case was inspected but has not been submitted yet. Customer due date: ${caseItem.customerDueDate}. Please submit your report at your earliest convenience. Thank you.`;
}

export function generateEmailSubject(caseItem) {
  const urgency = getUrgency(caseItem.customerDueDate);
  const prefix = urgency === 'overdue' ? 'OVERDUE - ' : urgency === 'due-soon' ? 'URGENT - ' : '';
  return `${prefix}Action Required: Submit Report for Control #${caseItem.controlNumber}`;
}

export function generateEmailBody(caseItem) {
  const frName = getFrName(caseItem.frAssigned);
  const urgency = getUrgency(caseItem.customerDueDate);
  const daysInfo = getUrgencyLabel(caseItem.customerDueDate);

  let urgencyBlock = '';
  if (urgency === 'overdue') {
    urgencyBlock = `\nIMPORTANT: This case is currently ${daysInfo}. Immediate action is required.\n`;
  } else if (urgency === 'due-soon') {
    urgencyBlock = `\nPlease note: There are only ${daysInfo} to submit this case.\n`;
  }

  return `Hello ${frName},

I hope this message finds you well. I am writing to follow up on the following case that has been inspected but not yet submitted:
${urgencyBlock}
Case Details:
- Control #: ${caseItem.controlNumber}
- Customer: ${caseItem.customerName}
- Survey Type: ${caseItem.surveyType}
- Location: ${caseItem.address}, ${caseItem.city}, ${caseItem.state} ${caseItem.zip}
- Date Ordered: ${caseItem.dateOrdered}
- Appointment Date: ${caseItem.appointmentDate}
- Date Returned: ${caseItem.dateReturned || 'N/A'}
- Customer Due Date: ${caseItem.customerDueDate}
- Field Status: ${caseItem.fieldStatus}${caseItem.lastNote ? `\n- Last Note: ${caseItem.lastNote}` : ''}

Please submit this report as soon as possible to meet the customer deadline.

If you are experiencing any issues or need assistance, please let me know immediately so we can resolve them.

Thank you for your prompt attention to this matter.

Best regards`;
}

export function generateBatchTextReminder(frName, cases) {
  const firstName = frName.split(' ')[0];
  const overdue = cases.filter((c) => getUrgency(c.customerDueDate) === 'overdue');
  const dueSoon = cases.filter((c) => getUrgency(c.customerDueDate) === 'due-soon');

  let urgencyNote = '';
  if (overdue.length > 0) {
    urgencyNote = ` ${overdue.length} ${overdue.length === 1 ? 'is' : 'are'} overdue.`;
  }
  if (dueSoon.length > 0) {
    urgencyNote += ` ${dueSoon.length} ${dueSoon.length === 1 ? 'is' : 'are'} due within 3 days.`;
  }

  const caseLines = cases
    .map((c) => `- #${c.controlNumber} (${c.surveyType}) at ${c.address}, ${c.city} - Due: ${c.customerDueDate}`)
    .join('\n');

  return `Hi ${firstName}, you have ${cases.length} case${cases.length !== 1 ? 's' : ''} that ${cases.length !== 1 ? 'were' : 'was'} inspected but not yet submitted.${urgencyNote}\n\n${caseLines}\n\nPlease submit these reports as soon as possible. Thank you.`;
}

export function generateBatchEmailBody(frName, cases) {
  const overdue = cases.filter((c) => getUrgency(c.customerDueDate) === 'overdue');
  const dueSoon = cases.filter((c) => getUrgency(c.customerDueDate) === 'due-soon');

  let urgencyBlock = '';
  if (overdue.length > 0) {
    urgencyBlock += `\nATTENTION: ${overdue.length} case${overdue.length !== 1 ? 's are' : ' is'} currently OVERDUE.\n`;
  }
  if (dueSoon.length > 0) {
    urgencyBlock += `${dueSoon.length} case${dueSoon.length !== 1 ? 's are' : ' is'} due within the next 3 days.\n`;
  }

  const caseTable = cases
    .map((c) => {
      const urg = getUrgency(c.customerDueDate);
      const flag = urg === 'overdue' ? ' [OVERDUE]' : urg === 'due-soon' ? ' [DUE SOON]' : '';
      return `  Control #${c.controlNumber}${flag}
    Customer: ${c.customerName}
    Type: ${c.surveyType}
    Location: ${c.address}, ${c.city}, ${c.state} ${c.zip}
    Due Date: ${c.customerDueDate}
    Date Returned: ${c.dateReturned || 'N/A'}${c.lastNote ? `\n    Note: ${c.lastNote}` : ''}`;
    })
    .join('\n\n');

  return `Hello ${frName},

I am following up on ${cases.length} case${cases.length !== 1 ? 's' : ''} assigned to you that ${cases.length !== 1 ? 'have' : 'has'} been inspected but not yet submitted.
${urgencyBlock}
${caseTable}

Please submit these reports at your earliest convenience to meet the customer deadlines.

If you are experiencing any issues or need assistance with any of these cases, please reach out immediately.

Thank you for your prompt attention.

Best regards`;
}

export function generateBatchEmailSubject(frName, cases) {
  const overdue = cases.filter((c) => getUrgency(c.customerDueDate) === 'overdue');
  const prefix = overdue.length > 0 ? 'OVERDUE - ' : '';
  return `${prefix}Action Required: ${cases.length} Inspected Case${cases.length !== 1 ? 's' : ''} Pending Submission`;
}
