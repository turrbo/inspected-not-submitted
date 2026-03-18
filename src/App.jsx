import { useState, useCallback, useMemo } from 'react';
import {
  Upload, FileSpreadsheet, Search, MessageSquare, Mail,
  Copy, Check, ChevronDown, ChevronRight, AlertTriangle,
  Clock, CheckCircle2, X, Users, BarChart3, Send, RefreshCw,
  Filter, Trash2
} from 'lucide-react';
import {
  parseExcel, getUrgency, getUrgencyLabel, getDaysUntilDue,
  getFrName, groupByFR,
  generateTextReminder, generateEmailSubject, generateEmailBody,
  generateBatchTextReminder, generateBatchEmailBody, generateBatchEmailSubject
} from './utils';
import './App.css';

function Toast({ message, type, onClose }) {
  return (
    <div className={`toast ${type}`} onClick={onClose}>
      {type === 'success' ? <CheckCircle2 size={18} /> : <Clock size={18} />}
      {message}
    </div>
  );
}

function CopyButton({ text }) {
  const [copied, setCopied] = useState(false);
  const handleCopy = async () => {
    await navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };
  return (
    <button className="btn-icon copy-btn" onClick={handleCopy} title="Copy to clipboard">
      {copied ? <Check size={14} color="var(--success)" /> : <Copy size={14} />}
    </button>
  );
}

function MessageModal({ caseItem, onClose }) {
  if (!caseItem) return null;
  const textMsg = generateTextReminder(caseItem);
  const emailSubject = generateEmailSubject(caseItem);
  const emailBody = generateEmailBody(caseItem);
  const urgency = getUrgency(caseItem.customerDueDate);

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <button className="btn-icon" style={{ position: 'absolute', top: 16, right: 16 }} onClick={onClose}>
          <X size={18} />
        </button>
        <h2>
          <MessageSquare size={22} />
          Messages for Control #{caseItem.controlNumber}
        </h2>
        <div style={{ display: 'flex', gap: 8, marginBottom: 20, flexWrap: 'wrap' }}>
          <span className="badge info">{caseItem.surveyType}</span>
          <span className={`badge ${urgency}`}>{getUrgencyLabel(caseItem.customerDueDate)}</span>
        </div>

        <div className="gen-section">
          <div className="message-label"><MessageSquare size={14} /> Text Reminder</div>
          <div className="message-box">
            <CopyButton text={textMsg} />
            <pre>{textMsg}</pre>
          </div>
        </div>

        <div className="gen-section">
          <div className="message-label"><Mail size={14} /> Email Subject</div>
          <div className="message-box">
            <CopyButton text={emailSubject} />
            <pre>{emailSubject}</pre>
          </div>
        </div>

        <div className="gen-section">
          <div className="message-label"><Mail size={14} /> Email Body</div>
          <div className="message-box">
            <CopyButton text={emailBody} />
            <pre>{emailBody}</pre>
          </div>
        </div>
      </div>
    </div>
  );
}

function BatchModal({ frName, cases, onClose }) {
  if (!frName || !cases) return null;
  const textMsg = generateBatchTextReminder(frName, cases);
  const emailSubject = generateBatchEmailSubject(frName, cases);
  const emailBody = generateBatchEmailBody(frName, cases);

  const overdue = cases.filter((c) => getUrgency(c.customerDueDate) === 'overdue').length;
  const dueSoon = cases.filter((c) => getUrgency(c.customerDueDate) === 'due-soon').length;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <button className="btn-icon" style={{ position: 'absolute', top: 16, right: 16 }} onClick={onClose}>
          <X size={18} />
        </button>
        <h2>
          <Users size={22} />
          Batch Messages for {frName}
        </h2>
        <div style={{ display: 'flex', gap: 8, marginBottom: 20, flexWrap: 'wrap' }}>
          <span className="badge info">{cases.length} case{cases.length !== 1 ? 's' : ''}</span>
          {overdue > 0 && <span className="badge overdue">{overdue} overdue</span>}
          {dueSoon > 0 && <span className="badge due-soon">{dueSoon} due soon</span>}
        </div>

        <div className="gen-section">
          <div className="message-label"><MessageSquare size={14} /> Batch Text Reminder</div>
          <div className="message-box">
            <CopyButton text={textMsg} />
            <pre>{textMsg}</pre>
          </div>
        </div>

        <div className="gen-section">
          <div className="message-label"><Mail size={14} /> Email Subject</div>
          <div className="message-box">
            <CopyButton text={emailSubject} />
            <pre>{emailSubject}</pre>
          </div>
        </div>

        <div className="gen-section">
          <div className="message-label"><Mail size={14} /> Email Body</div>
          <div className="message-box">
            <CopyButton text={emailBody} />
            <pre>{emailBody}</pre>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [cases, setCases] = useState([]);
  const [fileName, setFileName] = useState('');
  const [search, setSearch] = useState('');
  const [frFilter, setFrFilter] = useState('all');
  const [urgencyFilter, setUrgencyFilter] = useState('all');
  const [tab, setTab] = useState('table');
  const [selectedCase, setSelectedCase] = useState(null);
  const [batchFR, setBatchFR] = useState(null);
  const [toast, setToast] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [expandedFRs, setExpandedFRs] = useState({});
  const [selectedCases, setSelectedCases] = useState(new Set());
  const [hideReturned, setHideReturned] = useState(false);

  const returnedCount = useMemo(() => cases.filter((c) => c.dateReturned).length, [cases]);

  const removeReturnedCases = () => {
    const remaining = cases.filter((c) => !c.dateReturned);
    const removed = cases.length - remaining.length;
    setCases(remaining);
    setSelectedCases(new Set());
    showToast(`Removed ${removed} case${removed !== 1 ? 's' : ''} with a Date Returned`);
  };

  const showToast = (message, type = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 3000);
  };

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
      showToast('Please upload an Excel file (.xlsx, .xls, or .csv)', 'info');
      return;
    }
    try {
      const data = await parseExcel(file);
      setCases(data);
      setFileName(file.name);
      setSelectedCases(new Set());
      showToast(`Loaded ${data.length} cases from ${file.name}`);
    } catch (err) {
      showToast('Error parsing file: ' + err.message, 'info');
    }
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    handleFile(e.dataTransfer.files[0]);
  }, [handleFile]);

  const filtered = useMemo(() => {
    return cases.filter((c) => {
      if (hideReturned && c.dateReturned) return false;
      if (search) {
        const q = search.toLowerCase();
        const searchable = [c.controlNumber, c.frAssigned, c.customerName, c.address, c.city, c.state, c.zip, c.surveyType].join(' ').toLowerCase();
        if (!searchable.includes(q)) return false;
      }
      if (frFilter !== 'all' && getFrName(c.frAssigned) !== frFilter) return false;
      if (urgencyFilter !== 'all' && getUrgency(c.customerDueDate) !== urgencyFilter) return false;
      return true;
    });
  }, [cases, search, frFilter, urgencyFilter, hideReturned]);

  const frNames = useMemo(() => [...new Set(cases.map((c) => getFrName(c.frAssigned)))].sort(), [cases]);
  const frGroups = useMemo(() => groupByFR(filtered), [filtered]);

  const stats = useMemo(() => {
    const overdue = cases.filter((c) => getUrgency(c.customerDueDate) === 'overdue').length;
    const dueSoon = cases.filter((c) => getUrgency(c.customerDueDate) === 'due-soon').length;
    const onTrack = cases.filter((c) => getUrgency(c.customerDueDate) === 'on-track').length;
    return { total: cases.length, overdue, dueSoon, onTrack, frCount: frNames.length };
  }, [cases, frNames]);

  const toggleFR = (name) => {
    setExpandedFRs((prev) => ({ ...prev, [name]: !prev[name] }));
  };

  const toggleSelectCase = (controlNumber) => {
    setSelectedCases((prev) => {
      const next = new Set(prev);
      if (next.has(controlNumber)) next.delete(controlNumber);
      else next.add(controlNumber);
      return next;
    });
  };

  const selectAll = () => {
    if (selectedCases.size === filtered.length) {
      setSelectedCases(new Set());
    } else {
      setSelectedCases(new Set(filtered.map((c) => c.controlNumber)));
    }
  };

  const handleBatchSelected = () => {
    const selected = filtered.filter((c) => selectedCases.has(c.controlNumber));
    const groups = {};
    selected.forEach((c) => {
      const name = getFrName(c.frAssigned);
      if (!groups[name]) groups[name] = [];
      groups[name].push(c);
    });
    const entries = Object.entries(groups);
    if (entries.length === 1) {
      setBatchFR({ name: entries[0][0], cases: entries[0][1] });
    } else {
      setBatchFR({ name: 'Selected Field Reps', cases: selected });
    }
  };

  if (cases.length === 0) {
    return (
      <div className="app-container">
        <div style={{ maxWidth: 600, margin: '80px auto 0' }}>
          <div style={{ textAlign: 'center', marginBottom: 40 }}>
            <FileSpreadsheet size={48} color="var(--accent)" style={{ marginBottom: 16 }} />
            <h1 style={{ fontSize: 28, fontWeight: 800, marginBottom: 8, color: 'var(--heading)' }}>
              Inspected Not Submitted
            </h1>
            <p style={{ color: 'var(--text-light)', fontSize: 15 }}>
              Case Tracker & Reminder Generator
            </p>
          </div>
          <div
            className={`upload-zone ${dragOver ? 'drag-over' : ''}`}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={onDrop}
            onClick={() => document.getElementById('file-input').click()}
          >
            <Upload size={40} color="var(--text-light)" style={{ marginBottom: 16 }} />
            <h3 style={{ fontSize: 18, fontWeight: 600, marginBottom: 6, color: 'var(--heading)' }}>
              Drop your Excel file here
            </h3>
            <p style={{ color: 'var(--text)', fontSize: 14 }}>
              or click to browse -- supports .xlsx, .xls, .csv
            </p>
            <input
              id="file-input"
              type="file"
              accept=".xlsx,.xls,.csv"
              style={{ display: 'none' }}
              onChange={(e) => handleFile(e.target.files[0])}
            />
          </div>
          <div style={{ textAlign: 'center', marginTop: 32, color: 'var(--text-light)', fontSize: 13 }}>
            Upload your "Inspected Not Submitted" report to auto-generate text reminders and email drafts for each case.
          </div>
        </div>
        {toast && <Toast {...toast} onClose={() => setToast(null)} />}
      </div>
    );
  }

  return (
    <div className="app-container">
      <div className="header">
        <div>
          <h1>
            <FileSpreadsheet size={24} color="var(--accent)" />
            Inspected Not Submitted
          </h1>
          <div className="subtitle">
            {fileName} -- {cases.length} cases loaded
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <button
            className="btn btn-secondary"
            onClick={() => {
              setCases([]);
              setFileName('');
              setSelectedCases(new Set());
            }}
          >
            <RefreshCw size={14} /> New Upload
          </button>
        </div>
      </div>

      {/* Stats */}
      <div className="stats-grid">
        <div className="stat-card">
          <div className="label">Total Cases</div>
          <div className="value">{stats.total}</div>
          <div className="sub">{stats.frCount} field reps assigned</div>
        </div>
        <div className="stat-card">
          <div className="label">Overdue</div>
          <div className="value" style={{ color: 'var(--danger)' }}>{stats.overdue}</div>
          <div className="sub">Past customer due date</div>
        </div>
        <div className="stat-card">
          <div className="label">Due Soon</div>
          <div className="value" style={{ color: 'var(--warning)' }}>{stats.dueSoon}</div>
          <div className="sub">Within 3 days</div>
        </div>
        <div className="stat-card">
          <div className="label">On Track</div>
          <div className="value" style={{ color: 'var(--success)' }}>{stats.onTrack}</div>
          <div className="sub">More than 3 days remaining</div>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12, marginBottom: 16 }}>
        <div className="tabs">
          <button className={`tab ${tab === 'table' ? 'active' : ''}`} onClick={() => setTab('table')}>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              <BarChart3 size={14} /> All Cases
            </span>
          </button>
          <button className={`tab ${tab === 'byFR' ? 'active' : ''}`} onClick={() => setTab('byFR')}>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              <Users size={14} /> By Field Rep
            </span>
          </button>
        </div>
      </div>

      {/* Toolbar */}
      <div className="toolbar">
        <div className="search-bar">
          <Search size={16} color="var(--text-light)" />
          <input
            placeholder="Search cases..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>
        <select value={frFilter} onChange={(e) => setFrFilter(e.target.value)}>
          <option value="all">All Field Reps</option>
          {frNames.map((name) => <option key={name} value={name}>{name}</option>)}
        </select>
        <select value={urgencyFilter} onChange={(e) => setUrgencyFilter(e.target.value)}>
          <option value="all">All Urgency</option>
          <option value="overdue">Overdue</option>
          <option value="due-soon">Due Soon</option>
          <option value="on-track">On Track</option>
        </select>
        <button
          className={`btn btn-sm ${hideReturned ? 'btn-primary' : 'btn-secondary'}`}
          onClick={() => setHideReturned(!hideReturned)}
          title="Toggle visibility of cases that have a Date Returned"
        >
          <Filter size={13} />
          {hideReturned ? 'Showing: No Returned' : 'Hide Returned'}
        </button>
        {returnedCount > 0 && (
          <button
            className="btn btn-sm btn-danger"
            onClick={removeReturnedCases}
            title="Permanently remove cases with a Date Returned from the list"
          >
            <Trash2 size={13} />
            Remove Returned ({returnedCount})
          </button>
        )}
        <div style={{ color: 'var(--text-light)', fontSize: 12, marginLeft: 'auto' }}>
          {filtered.length} of {cases.length} cases
        </div>
      </div>

      {/* Batch bar */}
      {selectedCases.size > 0 && (
        <div className="batch-bar">
          <span className="count">{selectedCases.size}</span> case{selectedCases.size !== 1 ? 's' : ''} selected
          <button className="btn btn-sm" style={{ background: 'rgba(255,255,255,0.2)', color: '#fff', marginLeft: 'auto' }} onClick={handleBatchSelected}>
            <Send size={13} /> Generate Batch Messages
          </button>
          <button className="btn btn-sm" style={{ background: 'rgba(255,255,255,0.15)', color: '#fff' }} onClick={() => setSelectedCases(new Set())}>
            Clear
          </button>
        </div>
      )}

      {/* Table View */}
      {tab === 'table' && (
        <div className="table-wrapper" style={{ maxHeight: '60vh', overflowY: 'auto' }}>
          <table>
            <thead>
              <tr>
                <th style={{ width: 40 }}>
                  <div
                    className={`checkbox ${selectedCases.size === filtered.length && filtered.length > 0 ? 'checked' : ''}`}
                    onClick={selectAll}
                  >
                    {selectedCases.size === filtered.length && filtered.length > 0 && <Check size={12} color="#fff" />}
                  </div>
                </th>
                <th>Control #</th>
                <th>FR Assigned</th>
                <th>Customer</th>
                <th>Survey Type</th>
                <th>Appt Date</th>
                <th>Returned</th>
                <th>Due Date</th>
                <th>Urgency</th>
                <th>Location</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((c) => {
                const urg = getUrgency(c.customerDueDate);
                return (
                  <tr key={c.controlNumber} className={selectedCases.has(c.controlNumber) ? 'selected' : ''}>
                    <td>
                      <div
                        className={`checkbox ${selectedCases.has(c.controlNumber) ? 'checked' : ''}`}
                        onClick={() => toggleSelectCase(c.controlNumber)}
                      >
                        {selectedCases.has(c.controlNumber) && <Check size={12} color="#fff" />}
                      </div>
                    </td>
                    <td style={{ fontWeight: 600, fontFamily: 'monospace', color: 'var(--heading)' }}>{c.controlNumber}</td>
                    <td>{getFrName(c.frAssigned)}</td>
                    <td title={c.customerName} style={{ maxWidth: 180 }}>{c.customerName}</td>
                    <td title={c.surveyType} style={{ maxWidth: 160 }}>{c.surveyType}</td>
                    <td>{c.appointmentDate}</td>
                    <td>{c.dateReturned || '--'}</td>
                    <td>{c.customerDueDate}</td>
                    <td>
                      <span className={`badge ${urg}`}>
                        {urg === 'overdue' && <AlertTriangle size={11} />}
                        {urg === 'due-soon' && <Clock size={11} />}
                        {urg === 'on-track' && <CheckCircle2 size={11} />}
                        {getUrgencyLabel(c.customerDueDate)}
                      </span>
                    </td>
                    <td title={`${c.address}, ${c.city}, ${c.state}`} style={{ maxWidth: 180 }}>
                      {c.city}, {c.state}
                    </td>
                    <td>
                      <div style={{ display: 'flex', gap: 4 }}>
                        <button className="btn-icon" title="Generate text reminder" onClick={() => setSelectedCase(c)}>
                          <MessageSquare size={14} />
                        </button>
                        <button className="btn-icon" title="Generate email" onClick={() => setSelectedCase(c)}>
                          <Mail size={14} />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {filtered.length === 0 && (
            <div style={{ textAlign: 'center', padding: 40, color: 'var(--text-light)' }}>
              No cases match the current filters.
            </div>
          )}
        </div>
      )}

      {/* By FR View */}
      {tab === 'byFR' && (
        <div>
          {Object.entries(frGroups)
            .sort(([, a], [, b]) => b.length - a.length)
            .map(([name, frCases]) => {
              const overdue = frCases.filter((c) => getUrgency(c.customerDueDate) === 'overdue').length;
              const dueSoon = frCases.filter((c) => getUrgency(c.customerDueDate) === 'due-soon').length;
              const expanded = expandedFRs[name] !== false;

              return (
                <div className="fr-group" key={name}>
                  <div className="fr-group-header" onClick={() => toggleFR(name)}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                      {expanded ? <ChevronDown size={18} /> : <ChevronRight size={18} />}
                      <div>
                        <div style={{ fontWeight: 600, fontSize: 15, color: 'var(--heading)' }}>{name}</div>
                        <div style={{ fontSize: 12, color: 'var(--text-light)' }}>
                          {frCases.length} case{frCases.length !== 1 ? 's' : ''}
                        </div>
                      </div>
                      <div style={{ display: 'flex', gap: 6 }}>
                        {overdue > 0 && <span className="badge overdue">{overdue} overdue</span>}
                        {dueSoon > 0 && <span className="badge due-soon">{dueSoon} due soon</span>}
                      </div>
                    </div>
                    <button
                      className="btn btn-sm btn-primary"
                      onClick={(e) => {
                        e.stopPropagation();
                        setBatchFR({ name, cases: frCases });
                      }}
                    >
                      <Send size={13} /> Generate Messages
                    </button>
                  </div>
                  {expanded && (
                    <div className="fr-group-content">
                      {frCases.map((c) => {
                        const urg = getUrgency(c.customerDueDate);
                        return (
                          <div className="fr-case-item" key={c.controlNumber}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 0 }}>
                              <span style={{ fontFamily: 'monospace', fontWeight: 600, flexShrink: 0, color: 'var(--heading)' }}>
                                #{c.controlNumber}
                              </span>
                              <span style={{ color: 'var(--text-light)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                {c.surveyType}
                              </span>
                              <span style={{ color: 'var(--text-light)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                {c.city}, {c.state}
                              </span>
                            </div>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                              <span className={`badge ${urg}`}>
                                {getUrgencyLabel(c.customerDueDate)}
                              </span>
                              <button className="btn-icon" onClick={() => setSelectedCase(c)} title="Generate messages">
                                <MessageSquare size={14} />
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            })}
        </div>
      )}

      {/* Modals */}
      {selectedCase && <MessageModal caseItem={selectedCase} onClose={() => setSelectedCase(null)} />}
      {batchFR && <BatchModal frName={batchFR.name} cases={batchFR.cases} onClose={() => setBatchFR(null)} />}
      {toast && <Toast {...toast} onClose={() => setToast(null)} />}
    </div>
  );
}
