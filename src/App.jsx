import { useState, useRef, useEffect } from "react";
import * as mammoth from "mammoth";
import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

const C = {
  bg:'#0B1120', surface:'#111827', card:'#1F2937', border:'#374151',
  accent:'#F97316', text:'#F9FAFB', muted:'#9CA3AF', faint:'#6B7280',
  success:'#10B981', warn:'#F59E0B', error:'#EF4444', blue:'#3B82F6',
};
const st = {
  app:{ display:'flex', height:'100vh', background:C.bg, color:C.text, fontFamily:"'Inter',sans-serif", overflow:'hidden' },
  sidebar:{ width:220, background:C.surface, borderRight:'1px solid #1F2937', display:'flex', flexDirection:'column', flexShrink:0 },
  main:{ flex:1, overflow:'auto', display:'flex', flexDirection:'column' },
  header:{ padding:'18px 28px 14px', borderBottom:'1px solid #1F2937', background:C.surface, flexShrink:0 },
  content:{ padding:28, flex:1 },
  card:{ background:C.card, borderRadius:12, border:`1px solid ${C.border}`, padding:20 },
  input:{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:8, padding:'10px 14px', color:C.text, fontSize:14, width:'100%', outline:'none', boxSizing:'border-box' },
  textarea:{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:8, padding:'10px 14px', color:C.text, fontSize:14, width:'100%', outline:'none', boxSizing:'border-box', resize:'vertical' },
  label:{ fontSize:11, fontWeight:700, color:C.muted, marginBottom:6, display:'block', textTransform:'uppercase', letterSpacing:'0.6px' },
  tab:{ padding:'10px 16px', cursor:'pointer', fontSize:13, fontWeight:500, borderBottom:'2px solid transparent', whiteSpace:'nowrap' },
};
const btn = (v='primary',extra={}) => ({
  display:'inline-flex', alignItems:'center', justifyContent:'center', gap:7,
  padding: v==='sm'?'5px 13px':'10px 20px', borderRadius:8,
  fontSize: v==='sm'?12:14, fontWeight:600, cursor:'pointer',
  border: v==='outline'?`1px solid ${C.border}`:'none',
  background: v==='primary'?C.accent: v==='ghost'?'#374151': v==='outline'?'transparent': v==='danger'?'#EF444420':'#374151',
  color: v==='outline'?C.muted: v==='danger'?C.error:'#fff',
  ...extra
});
const badge = (t) => ({
  display:'inline-flex', alignItems:'center', padding:'2px 9px', borderRadius:20, fontSize:11, fontWeight:600,
  background: t==='prospect'?'#1D4ED815':'#05966915',
  color: t==='prospect'?'#60A5FA':'#34D399',
  border:`1px solid ${t==='prospect'?'#1D4ED840':'#05966940'}`,
});
const statusBadge = (s) => {
  const map = { pending:['#F59E0B20','#F59E0B'], in_progress:['#3B82F620','#3B82F6'], complete:['#10B98120','#10B981'] };
  const [bg,col] = map[s]||map.pending;
  return { display:'inline-flex', alignItems:'center', padding:'2px 9px', borderRadius:20, fontSize:11, fontWeight:600, background:bg, color:col, border:`1px solid ${col}40` };
};
const navItem = (a) => ({
  display:'flex', alignItems:'center', gap:9, padding:'9px 12px', borderRadius:8,
  cursor:'pointer', fontSize:13, fontWeight:a?600:400,
  color:a?C.accent:C.muted, background:a?'#374151':'transparent', marginBottom:2,
});

const DECK_TYPES = [
  { id:'sales', label:'Sales Deck', icon:'🎯', desc:'Prospect — first pitch' },
  { id:'monthly', label:'Monthly Recap', icon:'📅', desc:'Active client — MoM performance' },
  { id:'qbr', label:'Quarterly QBR', icon:'📊', desc:'Active client — full quarter review' },
  { id:'annual', label:'Annual Recap', icon:'🏆', desc:'Active client — year in review' },
  { id:'custom', label:'Custom', icon:'✏️', desc:'Anything outside standard cadence' },
];

const REPORT_TYPES = [
  { id:'sb_search_term', label:'SB Search Term Report', cat:'Sponsored Brands', deckTypes:['all'] },
  { id:'sb_kw_placement', label:'SB Keyword Placement', cat:'Sponsored Brands', deckTypes:['all'] },
  { id:'sp_placement', label:'SP Placement Report', cat:'Sponsored Products', deckTypes:['all'] },
  { id:'sp_advertised', label:'SP Advertised Product', cat:'Sponsored Products', deckTypes:['all'] },
  { id:'sp_search_term', label:'SP Search Term Report', cat:'Sponsored Products', deckTypes:['all'] },
  { id:'sqr', label:'SQR (Search Query Report)', cat:'Sponsored Products', deckTypes:['all'] },
  { id:'beta_product', label:'Beta Report by Product', cat:'DSP / Beta', deckTypes:['all'] },
  { id:'beta_target', label:'Beta Report by Target', cat:'DSP / Beta', deckTypes:['all'] },
  { id:'beta_campaign', label:'Beta Report by Campaign', cat:'DSP / Beta', deckTypes:['all'] },
  { id:'sc_sales_asin', label:'SC Total Sales by ASIN', cat:'Seller Central', deckTypes:['all'] },
  { id:'amc_impression_freq', label:'AMC: Impression Frequency', cat:'AMC', deckTypes:['all'] },
  { id:'amc_kw_purchase_path', label:'AMC: Keyword Purchase Path', cat:'AMC', deckTypes:['all'] },
  { id:'amc_ad_placement', label:'AMC: Ad Placement Effectiveness', cat:'AMC', deckTypes:['all'] },
  { id:'amc_time_to_conversion', label:'AMC: Time to Conversion', cat:'AMC', deckTypes:['all'] },
  { id:'ss_search', label:'SmartScout: Search Terms', cat:'SmartScout', deckTypes:['all'] },
  { id:'ss_market', label:'SmartScout: Market Share', cat:'SmartScout', deckTypes:['all'] },
  { id:'ss_products', label:'SmartScout: Products', cat:'SmartScout', deckTypes:['all'] },
  { id:'ss_brand', label:'SmartScout: Brand Monthly', cat:'SmartScout', deckTypes:['all'] },
  { id:'q1_last_year', label:'Q1 Last Year', cat:'Period Comparison', deckTypes:['qbr','sales','custom'] },
  { id:'q1_this_year', label:'Q1 This Year', cat:'Period Comparison', deckTypes:['qbr','sales','custom'] },
  { id:'mom_this_month', label:'MoM: This Month', cat:'Month over Month', deckTypes:['monthly','sales'] },
  { id:'mom_last_month', label:'MoM: Last Month', cat:'Month over Month', deckTypes:['monthly','sales'] },
  { id:'yoy_this_year', label:'YoY: This Year', cat:'Year over Year', deckTypes:['annual','sales'] },
  { id:'yoy_last_year', label:'YoY: Last Year', cat:'Year over Year', deckTypes:['annual','sales'] },
];

const CAT_COLORS = {
  'Sponsored Brands':'#3B82F6','Sponsored Products':'#8B5CF6','DSP / Beta':'#F97316',
  'Seller Central':'#10B981','AMC':'#EC4899','SmartScout':'#F59E0B',
  'Period Comparison':'#84CC16','Month over Month':'#06B6D4','Year over Year':'#A78BFA',
};

const BTR_SYSTEM_PROMPT = `You are acting as a senior retail media strategist at BTR Media. Analyze decks and create reusable style guides for future deck building.`;

const reportsForDeckType = (deckType) =>
  REPORT_TYPES.filter(r => r.deckTypes.includes('all') || r.deckTypes.includes(deckType));

export default function BTRDeckStudio() {
  const [screen, setScreen] = useState('clients');
  const [clients, setClients] = useState([]);
  const [loadingClients, setLoadingClients] = useState(true);
  const [selId, setSelId] = useState(null);
  const [activeTab, setActiveTab] = useState('overview');
  const [styleRefs, setStyleRefs] = useState([]);
  const [loadingStyleRefs, setLoadingStyleRefs] = useState(true);
  const [processingStyleRef, setProcessingStyleRef] = useState(false);
  const [clientReports, setClientReports] = useState([]);
  const [clientTranscripts, setClientTranscripts] = useState([]);
  const [clientDecks, setClientDecks] = useState([]);
  const [clientRequests, setClientRequests] = useState([]);
  const [clientNotes, setClientNotes] = useState([]);
  const [loadingClientData, setLoadingClientData] = useState(false);
  const [selectedRequest, setSelectedRequest] = useState(null);
  const [uploadingReportType, setUploadingReportType] = useState(null);
  const [uploadingDeck, setUploadingDeck] = useState(false);
  const [briefingCopied, setBriefingCopied] = useState(false);

  const reportFileRef = useRef(null);
  const transcriptFileRef = useRef(null);
  const deckFileRef = useRef(null);
  const styleRefFileRef = useRef(null);
  const pendingUploadRef = useRef(null);

  useEffect(() => { loadClients(); loadStyleRefs(); }, []);

  const loadClients = async () => {
    setLoadingClients(true);
    const { data } = await supabase.from('clients').select('*').order('created_at', { ascending:false });
    if (data) setClients(data);
    setLoadingClients(false);
  };

  const loadClientData = async (id) => {
    setLoadingClientData(true);
    const [rp, tr, dk, rq, nt] = await Promise.all([
      supabase.from('reports').select('*').eq('client_id',id).order('uploaded_at',{ascending:false}),
      supabase.from('transcripts').select('*').eq('client_id',id).order('uploaded_at',{ascending:false}),
      supabase.from('decks').select('*').eq('client_id',id).order('created_at',{ascending:false}),
      supabase.from('deck_requests').select('*').eq('client_id',id).order('created_at',{ascending:false}),
      supabase.from('context_notes').select('*').eq('client_id',id).order('created_at',{ascending:false}),
    ]);
    setClientReports(rp.data||[]);
    setClientTranscripts(tr.data||[]);
    setClientDecks(dk.data||[]);
    setClientRequests(rq.data||[]);
    setClientNotes(nt.data||[]);
    setLoadingClientData(false);
  };

  const loadStyleRefs = async () => {
    setLoadingStyleRefs(true);
    const { data } = await supabase.from('style_references').select('*').order('created_at',{ascending:false});
    if (data) setStyleRefs(data);
    setLoadingStyleRefs(false);
  };

  const client = clients.find(c => c.id === selId);

  const openProfile = (id) => {
    setSelId(id);
    setActiveTab('overview');
    setScreen('profile');
    loadClientData(id);
  };

  // ── FILE UPLOADS ─────────────────────────────────────────────────────────────

  const uploadReportFile = async (file, reportTypeId) => {
    const rtype = REPORT_TYPES.find(r => r.id === reportTypeId);
    const reportId = crypto.randomUUID();
    const ext = file.name.split('.').pop();
    const storagePath = `${selId}/reports/${reportId}.${ext}`;
    const { error: storageErr } = await supabase.storage.from('client-files').upload(storagePath, file, { upsert:true });
    if (storageErr) { console.error(storageErr); return null; }
    const { data } = await supabase.from('reports').insert([{
      id: reportId,
      client_id: selId,
      report_type: reportTypeId,
      report_label: rtype.label,
      report_category: rtype.cat,
      deck_types: rtype.deckTypes,
      file_name: file.name,
      storage_path: storagePath,
    }]).select().single();
    if (data) setClientReports(prev => [data, ...prev]);
    return data;
  };

  const handleReportFile = async (e) => {
    const file = e.target.files[0]; e.target.value = '';
    if (!file || !pendingUploadRef.current) return;
    setUploadingReportType(pendingUploadRef.current);
    await uploadReportFile(file, pendingUploadRef.current);
    setUploadingReportType(null);
    pendingUploadRef.current = null;
  };

  const triggerReportUpload = (typeId) => {
    pendingUploadRef.current = typeId;
    reportFileRef.current?.click();
  };

  const uploadTranscript = async (file) => {
    let content = '';
    if (file.name.match(/\.docx?$/i)) {
      const buf = await file.arrayBuffer();
      const r = await mammoth.extractRawText({ arrayBuffer:buf });
      content = r.value;
    } else {
      content = await file.text();
    }
    const { data } = await supabase.from('transcripts').insert([{
      client_id: selId,
      transcript_type: 'discovery',
      title: file.name.replace(/\.[^.]+$/, ''),
      call_date: new Date().toISOString().split('T')[0],
      content,
    }]).select().single();
    if (data) setClientTranscripts(prev => [data, ...prev]);
  };

  const handleTranscriptFile = async (e) => {
    const file = e.target.files[0]; e.target.value = '';
    if (!file) return;
    await uploadTranscript(file);
  };

  const uploadDeckFile = async (file, requestId) => {
    setUploadingDeck(true);
    const deckId = crypto.randomUUID();
    const ext = file.name.split('.').pop();
    const storagePath = `${selId}/decks/${deckId}.${ext}`;
    const { error: storageErr } = await supabase.storage.from('client-files').upload(storagePath, file, { upsert:true });
    if (!storageErr) {
      const reqData = selectedRequest || clientRequests.find(r => r.id === requestId);
      const { data } = await supabase.from('decks').insert([{
        id: deckId,
        client_id: selId,
        deck_request_id: requestId || null,
        deck_type: reqData?.deck_type || 'custom',
        title: file.name.replace(/\.[^.]+$/, ''),
        storage_path: storagePath,
        file_name: file.name,
      }]).select().single();
      if (data) {
        setClientDecks(prev => [data, ...prev]);
        if (requestId) {
          await supabase.from('deck_requests').update({ status:'complete' }).eq('id', requestId);
          setClientRequests(prev => prev.map(r => r.id===requestId ? {...r, status:'complete'} : r));
          if (selectedRequest?.id === requestId) setSelectedRequest(prev => ({...prev, status:'complete'}));
        }
      }
    }
    setUploadingDeck(false);
  };

  const downloadReport = async (report) => {
    const { data } = await supabase.storage.from('client-files').createSignedUrl(report.storage_path, 3600);
    if (data?.signedUrl) { const a=document.createElement('a'); a.href=data.signedUrl; a.download=report.file_name; a.click(); }
  };

  const downloadDeck = async (deck) => {
    const { data } = await supabase.storage.from('client-files').createSignedUrl(deck.storage_path, 3600);
    if (data?.signedUrl) { const a=document.createElement('a'); a.href=data.signedUrl; a.download=deck.file_name; a.click(); }
  };

  const deleteReport = async (report) => {
    await supabase.storage.from('client-files').remove([report.storage_path]);
    await supabase.from('reports').delete().eq('id', report.id);
    setClientReports(prev => prev.filter(r => r.id !== report.id));
  };

  const deleteTranscript = async (id) => {
    await supabase.from('transcripts').delete().eq('id', id);
    setClientTranscripts(prev => prev.filter(t => t.id !== id));
  };

  // ── BRIEFING GENERATION ──────────────────────────────────────────────────────

  const generateBriefing = async (request) => {
    const attachedReports = clientReports.filter(r => (request.report_ids||[]).includes(r.id));
    const attachedTranscripts = clientTranscripts.filter(t => (request.transcript_ids||[]).includes(t.id));
    const deckType = DECK_TYPES.find(d => d.id === request.deck_type);
    const pastDecks = clientDecks.slice(0, 5);

    const briefing = `===========================
BTR MEDIA — DECK BUILD REQUEST
===========================
Client: ${client?.name} (${client?.type === 'prospect' ? 'Prospect' : 'Active Client'})
Deck Type: ${deckType?.label || request.deck_type}
Due Date: ${request.due_date || 'Not specified'}
Requested By: ${request.requested_by || 'Not specified'}
Prepared: ${new Date().toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'})}

SPECIAL INSTRUCTIONS
--------------------
${request.special_instructions || 'None'}

CLIENT PROFILE
--------------
Amazon Categories: ${client?.categories || 'Not specified'}
Current Goals: ${client?.goals || 'Not specified'}

${attachedTranscripts.length > 0 ? `TRANSCRIPTS (${attachedTranscripts.length} included — full content below)
---------------------------------------------
${attachedTranscripts.map(t => `--- ${t.transcript_type?.toUpperCase()} | ${t.call_date ? new Date(t.call_date).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}) : 'Date unknown'} ---
${t.content}`).join('\n\n')}` : 'TRANSCRIPTS: None attached to this request.'}

REPORTS (${attachedReports.length} included — download files below and upload to this Claude Project)
----------------------------------------------------------------------
${attachedReports.length > 0 ? attachedReports.map(r => `• ${r.report_label}${r.date_range ? ` (${r.date_range})` : ''} — ${r.file_name}`).join('\n') : 'No reports attached to this request.'}

PAST DECKS FOR THIS CLIENT (for context and learning)
------------------------------------------------------
${pastDecks.length > 0 ? pastDecks.map(d => `• ${d.title} — ${new Date(d.created_at).toLocaleDateString('en-US',{month:'short',year:'numeric'})} (${d.deck_type})`).join('\n') : 'No past decks on file yet.'}

INSTRUCTIONS FOR CLAUDE
-----------------------
You are acting as a senior retail media strategist at BTR Media. All uploaded report files for this client are attached to this conversation. Review every report in full — do not sample or summarize without reading all the data. Build a ${deckType?.label || request.deck_type} for ${client?.name} following BTR Media's narrative audit approach:

1. What We Heard — mirror back what the client told us
2. Account Snapshot — current state with data
3. The Problems — specific issues with exact evidence and dollar impact
4. Cost of Staying Here — revenue left on the table
5. The BTR Strategy — solutions mapped to each problem
6. What Becomes Possible — growth scenarios
7. 90-Day Roadmap — specific and actionable
8. Next Steps — clear actions

Rules: Use only numbers that appear in the uploaded report files. Never fabricate statistics. Every claim must be traceable to the data.`;

    await navigator.clipboard.writeText(briefing);
    setBriefingCopied(true);
    setTimeout(() => setBriefingCopied(false), 3000);
  };

  // ── STYLE LIBRARY ────────────────────────────────────────────────────────────

  const processStyleRef = async (file) => {
    setProcessingStyleRef(true);
    try {
      const base64 = await new Promise(res => { const r=new FileReader(); r.onload=e=>res(e.target.result.split(',')[1]); r.readAsDataURL(file); });
      const resp = await fetch('/api/ai', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({ model:'claude-sonnet-4-20250514', max_tokens:2000, system:BTR_SYSTEM_PROMPT,
          messages:[{ role:'user', content:[
            { type:'document', source:{ type:'base64', media_type:'application/pdf', data:base64 } },
            { type:'text', text:'Analyze this BTR Media deck and create a reusable style guide covering: narrative structure, content depth, data presentation, problem framing, solution framing, language and tone, slide structure, and storytelling approach. Output 400-600 words focused on HOW things are structured, not the specific brand content.' }
          ]}]
        })
      });
      const d = await resp.json();
      const summary = d.content?.[0]?.text||'';
      const name = file.name.replace(/\.pdf$/i,'').replace(/_/g,' ');
      const { data } = await supabase.from('style_references').insert([{ name, file_name:file.name, style_summary:summary, active:true }]).select().single();
      if (data) setStyleRefs(prev=>[data,...prev]);
    } catch(err) { console.error(err); }
    setProcessingStyleRef(false);
  };

  const toggleStyleRef = async (ref) => {
    const { data } = await supabase.from('style_references').update({ active:!ref.active }).eq('id',ref.id).select().single();
    if (data) setStyleRefs(prev=>prev.map(r=>r.id===ref.id?data:r));
  };

  const deleteStyleRef = async (id) => {
    await supabase.from('style_references').delete().eq('id',id);
    setStyleRefs(prev=>prev.filter(r=>r.id!==id));
  };

  // ── SCREENS ──────────────────────────────────────────────────────────────────

  const Clients = () => (
    <div>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:24 }}>
        <div style={{ fontSize:13, color:C.muted }}>{loadingClients?'Loading...':`${clients.length} clients · ${clients.filter(c=>c.type==='prospect').length} prospects · ${clients.filter(c=>c.type==='active').length} active`}</div>
        <button style={btn('primary')} onClick={()=>setScreen('new-client')}>+ New Client</button>
      </div>
      {loadingClients?<div style={{ textAlign:'center', padding:'60px 0', color:C.faint }}>Loading…</div>:(
        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(265px,1fr))', gap:16 }}>
          {clients.map(c=>(
            <div key={c.id} style={{ ...st.card, cursor:'pointer' }} onClick={()=>openProfile(c.id)}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:10 }}>
                <div style={{ fontSize:15, fontWeight:700, flex:1, marginRight:8 }}>{c.name}</div>
                <span style={badge(c.type)}>{c.type==='prospect'?'Prospect':'Active'}</span>
              </div>
              <div style={{ fontSize:13, color:C.muted, marginBottom:6 }}>👤 {c.contact||'—'}</div>
              <div style={{ fontSize:12, color:C.faint, marginBottom:14, lineHeight:1.5 }}>{c.categories||'—'}</div>
              <div style={{ borderTop:`1px solid ${C.border}`, paddingTop:12, display:'flex', justifyContent:'space-between', fontSize:12 }}>
                <span style={{ color:C.faint }}>Open client file →</span>
                <span style={{ color:C.accent, fontWeight:600 }}>→</span>
              </div>
            </div>
          ))}
          <div style={{ ...st.card, border:`1px dashed ${C.border}`, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:8, color:C.faint, fontSize:14, minHeight:150 }} onClick={()=>setScreen('new-client')}>
            <span style={{ fontSize:22 }}>＋</span> New Client
          </div>
        </div>
      )}
    </div>
  );

  const NewClient = () => {
    const [f,setF]=useState({ name:'', type:'prospect', contact:'', email:'', categories:'', goals:'' });
    const [saving,setSaving]=useState(false);
    const save=async()=>{ if(!f.name.trim()) return; setSaving(true); const{data}=await supabase.from('clients').insert([f]).select().single(); if(data){setClients(p=>[data,...p]); openProfile(data.id);} setSaving(false); };
    return (
      <div style={{ maxWidth:580 }}>
        <button style={{ ...btn('outline'), marginBottom:18 }} onClick={()=>setScreen('clients')}>← Back</button>
        <div style={st.card}>
          <div style={{ fontSize:16, fontWeight:700, marginBottom:20 }}>New Client Profile</div>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14 }}>
            <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Company Name *</label><input style={st.input} value={f.name} onChange={e=>setF(p=>({...p,name:e.target.value}))} placeholder="Brand name" /></div>
            <div><label style={st.label}>Client Type</label><select style={st.input} value={f.type} onChange={e=>setF(p=>({...p,type:e.target.value}))}><option value="prospect">Prospect</option><option value="active">Active Client</option></select></div>
            <div><label style={st.label}>Contact Name</label><input style={st.input} value={f.contact} onChange={e=>setF(p=>({...p,contact:e.target.value}))} placeholder="Name" /></div>
            <div><label style={st.label}>Email</label><input style={st.input} value={f.email} onChange={e=>setF(p=>({...p,email:e.target.value}))} placeholder="email@brand.com" /></div>
            <div><label style={st.label}>Amazon Categories</label><input style={st.input} value={f.categories} onChange={e=>setF(p=>({...p,categories:e.target.value}))} placeholder="Health, Beauty…" /></div>
            <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Client Goals</label><textarea style={{ ...st.textarea, minHeight:80 }} value={f.goals} onChange={e=>setF(p=>({...p,goals:e.target.value}))} placeholder="What are they trying to achieve on Amazon?" /></div>
          </div>
          <div style={{ display:'flex', gap:10, marginTop:18 }}>
            <button style={{ ...btn('primary'), opacity:(!f.name||saving)?0.5:1 }} onClick={save} disabled={!f.name||saving}>{saving?'Saving…':'Create Profile'}</button>
            <button style={btn('ghost')} onClick={()=>setScreen('clients')}>Cancel</button>
          </div>
        </div>
      </div>
    );
  };

  const Profile = () => {
    if (!client) return null;
    const tabs = [
      { id:'overview', label:'Overview' },
      { id:'reports', label:`Reports${clientReports.length>0?` (${clientReports.length})`:''}` },
      { id:'transcripts', label:`Transcripts${clientTranscripts.length>0?` (${clientTranscripts.length})`:''}` },
      { id:'requests', label:`Deck Requests${clientRequests.length>0?` (${clientRequests.length})`:''}` },
      { id:'decks', label:`Decks${clientDecks.length>0?` (${clientDecks.length})`:''}` },
      { id:'notes', label:'Notes' },
    ];
    return (
      <div>
        <button style={{ ...btn('outline'), marginBottom:18 }} onClick={()=>setScreen('clients')}>← All Clients</button>
        <div style={{ ...st.card, marginBottom:0, borderRadius:'12px 12px 0 0', borderBottom:`1px solid ${C.border}` }}>
          <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start' }}>
            <div>
              <div style={{ fontSize:22, fontWeight:800, marginBottom:6 }}>{client.name}</div>
              <div style={{ display:'flex', gap:8, alignItems:'center', flexWrap:'wrap' }}>
                <span style={badge(client.type)}>{client.type==='prospect'?'Prospect':'Active Client'}</span>
                {client.contact&&<span style={{ fontSize:12, color:C.faint }}>👤 {client.contact}</span>}
                {client.categories&&<span style={{ fontSize:12, color:C.faint }}>📦 {client.categories}</span>}
              </div>
            </div>
            <button style={btn('primary')} onClick={()=>setScreen('new-request')}>+ New Deck Request</button>
          </div>
        </div>
        <div style={{ background:C.card, borderLeft:`1px solid ${C.border}`, borderRight:`1px solid ${C.border}`, display:'flex', overflowX:'auto', borderBottom:`1px solid ${C.border}` }}>
          {tabs.map(t=>(
            <div key={t.id} style={{ ...st.tab, color:activeTab===t.id?C.accent:C.muted, borderBottomColor:activeTab===t.id?C.accent:'transparent', fontWeight:activeTab===t.id?700:500 }} onClick={()=>setActiveTab(t.id)}>{t.label}</div>
          ))}
        </div>
        <div style={{ ...st.card, borderRadius:'0 0 12px 12px', borderTop:'none' }}>
          {loadingClientData&&activeTab!=='overview'?<div style={{ textAlign:'center', padding:'40px', color:C.faint }}>Loading…</div>:null}
          {activeTab==='overview'&&<OverviewTab />}
          {activeTab==='reports'&&!loadingClientData&&<ReportsTab />}
          {activeTab==='transcripts'&&!loadingClientData&&<TranscriptsTab />}
          {activeTab==='requests'&&!loadingClientData&&<RequestsTab />}
          {activeTab==='decks'&&!loadingClientData&&<DecksTab />}
          {activeTab==='notes'&&!loadingClientData&&<NotesTab />}
        </div>
      </div>
    );
  };

  const OverviewTab = () => {
    const [editing,setEditing]=useState(false);
    const [ef,setEf]=useState({ name:client.name, type:client.type, contact:client.contact||'', email:client.email||'', categories:client.categories||'', goals:client.goals||'' });
    const [saving,setSaving]=useState(false);
    const save=async()=>{ setSaving(true); const{data}=await supabase.from('clients').update(ef).eq('id',selId).select().single(); if(data) setClients(prev=>prev.map(c=>c.id===selId?data:c)); setSaving(false); setEditing(false); };
    return (
      <div>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:20 }}>
          <div style={{ fontSize:15, fontWeight:700 }}>Client Profile</div>
          {!editing&&<button style={btn('ghost')} onClick={()=>setEditing(true)}>✏️ Edit</button>}
        </div>
        {editing?(
          <div>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12, marginBottom:16 }}>
              <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Company Name</label><input style={st.input} value={ef.name} onChange={e=>setEf(p=>({...p,name:e.target.value}))} /></div>
              <div><label style={st.label}>Type</label><select style={st.input} value={ef.type} onChange={e=>setEf(p=>({...p,type:e.target.value}))}><option value="prospect">Prospect</option><option value="active">Active Client</option></select></div>
              <div><label style={st.label}>Contact</label><input style={st.input} value={ef.contact} onChange={e=>setEf(p=>({...p,contact:e.target.value}))} /></div>
              <div><label style={st.label}>Email</label><input style={st.input} value={ef.email} onChange={e=>setEf(p=>({...p,email:e.target.value}))} /></div>
              <div><label style={st.label}>Categories</label><input style={st.input} value={ef.categories} onChange={e=>setEf(p=>({...p,categories:e.target.value}))} /></div>
              <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Goals</label><textarea style={{ ...st.textarea, minHeight:100 }} value={ef.goals} onChange={e=>setEf(p=>({...p,goals:e.target.value}))} /></div>
            </div>
            <div style={{ display:'flex', gap:10 }}>
              <button style={{ ...btn('primary'), opacity:saving?0.5:1 }} onClick={save} disabled={saving}>{saving?'Saving…':'Save'}</button>
              <button style={btn('ghost')} onClick={()=>setEditing(false)}>Cancel</button>
            </div>
          </div>
        ):(
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12 }}>
            {[['Contact',client.contact||'—'],['Email',client.email||'—'],['Categories',client.categories||'—']].map(([l,v])=>(
              <div key={l} style={{ background:C.surface, borderRadius:8, padding:'10px 14px' }}>
                <div style={st.label}>{l}</div>
                <div style={{ fontSize:13, color:C.text }}>{v}</div>
              </div>
            ))}
            <div style={{ background:C.surface, borderRadius:8, padding:'10px 14px', gridColumn:'1/-1' }}>
              <div style={st.label}>Goals</div>
              <div style={{ fontSize:13, color:C.text, lineHeight:1.7, whiteSpace:'pre-wrap' }}>{client.goals||'No goals set.'}</div>
            </div>
          </div>
        )}
        <div style={{ marginTop:20, display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:12 }}>
          {[['Reports',clientReports.length,'reports'],['Transcripts',clientTranscripts.length,'transcripts'],['Deck Requests',clientRequests.length,'requests'],['Decks',clientDecks.length,'decks']].map(([l,n,tab])=>(
            <div key={l} style={{ background:C.surface, borderRadius:8, padding:'14px', textAlign:'center', cursor:'pointer' }} onClick={()=>setActiveTab(tab)}>
              <div style={{ fontSize:24, fontWeight:800, color:C.accent }}>{n}</div>
              <div style={{ fontSize:12, color:C.muted, marginTop:2 }}>{l}</div>
            </div>
          ))}
        </div>
      </div>
    );
  };

  const ReportsTab = () => {
    const catGroups = REPORT_TYPES.reduce((acc,r)=>{ (acc[r.cat]=acc[r.cat]||[]).push(r); return acc; },{});
    return (
      <div>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:20 }}>
          <div style={{ fontSize:15, fontWeight:700 }}>Reports</div>
          <div style={{ fontSize:12, color:C.muted }}>{clientReports.length} file{clientReports.length!==1?'s':''} stored</div>
        </div>
        <input ref={reportFileRef} type="file" accept=".csv,.xlsx,.xls,.txt,.tsv" style={{ display:'none' }} onChange={handleReportFile} />
        {Object.entries(catGroups).map(([cat,rpts])=>(
          <div key={cat} style={{ marginBottom:20 }}>
            <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:10 }}>
              <span style={{ width:9, height:9, borderRadius:'50%', background:CAT_COLORS[cat]||C.muted, display:'inline-block' }}></span>
              <span style={{ fontSize:12, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'0.6px' }}>{cat}</span>
            </div>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(260px,1fr))', gap:8 }}>
              {rpts.map(rt=>{
                const existing = clientReports.filter(r=>r.report_type===rt.id);
                const isUploading = uploadingReportType===rt.id;
                return (
                  <div key={rt.id} style={{ background:C.surface, border:`1px solid ${existing.length>0?C.success:C.border}`, borderRadius:10, padding:'12px 14px' }}>
                    <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:existing.length>0?8:0 }}>
                      <div style={{ fontSize:13, fontWeight:600, color:existing.length>0?C.success:C.text, flex:1 }}>{existing.length>0&&'✓ '}{rt.label}</div>
                      <button style={{ ...btn('sm'), background:'transparent', border:`1px solid ${C.border}`, color:C.muted, flexShrink:0, marginLeft:8, opacity:isUploading?0.5:1 }} onClick={()=>triggerReportUpload(rt.id)} disabled={isUploading}>
                        {isUploading?'…':'⬆'}
                      </button>
                    </div>
                    {existing.map(r=>(
                      <div key={r.id} style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginTop:4 }}>
                        <div style={{ fontSize:11, color:C.faint, flex:1, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.file_name}</div>
                        <div style={{ display:'flex', gap:4, flexShrink:0, marginLeft:8 }}>
                          <button style={{ ...btn('sm'), padding:'2px 8px', background:'transparent', border:`1px solid ${C.border}`, color:C.muted, fontSize:11 }} onClick={()=>downloadReport(r)}>↓</button>
                          <button style={{ ...btn('sm'), padding:'2px 8px', background:'transparent', border:'1px solid #EF444430', color:C.error, fontSize:11 }} onClick={()=>deleteReport(r)}>✕</button>
                        </div>
                      </div>
                    ))}
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </div>
    );
  };

  const TranscriptsTab = () => {
    const TYPES = ['discovery','sales','qbr','check_in','other'];
    const [editingId,setEditingId]=useState(null);
    const [editType,setEditType]=useState('');
    const [editTitle,setEditTitle]=useState('');
    return (
      <div>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:20 }}>
          <div style={{ fontSize:15, fontWeight:700 }}>Transcripts</div>
          <button style={btn('primary')} onClick={()=>transcriptFileRef.current?.click()}>+ Upload Transcript</button>
        </div>
        <input ref={transcriptFileRef} type="file" accept=".txt,.doc,.docx" style={{ display:'none' }} onChange={handleTranscriptFile} />
        {clientTranscripts.length===0?(
          <div style={{ textAlign:'center', padding:'40px', color:C.faint }}>
            <div style={{ fontSize:32, marginBottom:12 }}>📝</div>
            <div style={{ fontSize:14, fontWeight:600, marginBottom:8 }}>No transcripts yet</div>
            <div style={{ fontSize:13, marginBottom:20 }}>Upload discovery calls, sales calls, QBRs, and check-ins.</div>
          </div>
        ):(
          clientTranscripts.map(t=>(
            <div key={t.id} style={{ ...st.card, marginBottom:10, padding:16 }}>
              {editingId===t.id?(
                <div style={{ display:'flex', gap:10, alignItems:'center' }}>
                  <select style={{ ...st.input, width:'auto' }} value={editType} onChange={e=>setEditType(e.target.value)}>
                    {TYPES.map(tt=><option key={tt} value={tt}>{tt.replace('_',' ')}</option>)}
                  </select>
                  <input style={{ ...st.input, flex:1 }} value={editTitle} onChange={e=>setEditTitle(e.target.value)} />
                  <button style={btn('primary')} onClick={async()=>{ await supabase.from('transcripts').update({ transcript_type:editType, title:editTitle }).eq('id',t.id); setClientTranscripts(prev=>prev.map(tr=>tr.id===t.id?{...tr,transcript_type:editType,title:editTitle}:tr)); setEditingId(null); }}>Save</button>
                  <button style={btn('ghost')} onClick={()=>setEditingId(null)}>✕</button>
                </div>
              ):(
                <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                  <div>
                    <div style={{ fontSize:14, fontWeight:600 }}>{t.title||'Untitled Transcript'}</div>
                    <div style={{ fontSize:12, color:C.muted, marginTop:3 }}>
                      <span style={{ background:'#374151', padding:'1px 8px', borderRadius:20, marginRight:8, fontSize:11 }}>{(t.transcript_type||'discovery').replace('_',' ')}</span>
                      {t.call_date&&new Date(t.call_date).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})}
                      {t.content&&<span style={{ marginLeft:8 }}>{t.content.split(/\s+/).length.toLocaleString()} words</span>}
                    </div>
                  </div>
                  <div style={{ display:'flex', gap:8 }}>
                    <button style={{ ...btn('sm'), background:'transparent', border:`1px solid ${C.border}`, color:C.muted }} onClick={()=>{ setEditingId(t.id); setEditType(t.transcript_type||'discovery'); setEditTitle(t.title||''); }}>Edit</button>
                    <button style={{ ...btn('sm'), background:'transparent', border:'1px solid #EF444430', color:C.error }} onClick={()=>deleteTranscript(t.id)}>Delete</button>
                  </div>
                </div>
              )}
            </div>
          ))
        )}
      </div>
    );
  };

  const RequestsTab = () => (
    <div>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:20 }}>
        <div style={{ fontSize:15, fontWeight:700 }}>Deck Requests</div>
        <button style={btn('primary')} onClick={()=>setScreen('new-request')}>+ New Request</button>
      </div>
      {clientRequests.length===0?(
        <div style={{ textAlign:'center', padding:'40px', color:C.faint }}>
          <div style={{ fontSize:32, marginBottom:12 }}>📋</div>
          <div style={{ fontSize:14, fontWeight:600, marginBottom:8 }}>No deck requests yet</div>
          <div style={{ fontSize:13, marginBottom:20 }}>Create a request to organize reports and generate a Claude briefing.</div>
          <button style={btn('primary')} onClick={()=>setScreen('new-request')}>Create First Request</button>
        </div>
      ):(
        clientRequests.map(req=>{
          const dtype = DECK_TYPES.find(d=>d.id===req.deck_type);
          return (
            <div key={req.id} style={{ ...st.card, marginBottom:10, padding:16, cursor:'pointer' }} onClick={()=>{ setSelectedRequest(req); setScreen('request-detail'); }}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start' }}>
                <div>
                  <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:6 }}>
                    <span style={{ fontSize:16 }}>{dtype?.icon||'📋'}</span>
                    <div style={{ fontSize:14, fontWeight:700 }}>{req.title||`${dtype?.label} — ${client?.name}`}</div>
                    <span style={statusBadge(req.status)}>{(req.status||'pending').replace('_',' ')}</span>
                  </div>
                  <div style={{ fontSize:12, color:C.faint }}>
                    {req.due_date&&<span>Due {new Date(req.due_date).toLocaleDateString('en-US',{month:'short',day:'numeric'})} · </span>}
                    {req.requested_by&&<span>Requested by {req.requested_by} · </span>}
                    <span>{(req.report_ids||[]).length} reports · {(req.transcript_ids||[]).length} transcripts</span>
                  </div>
                </div>
                <span style={{ color:C.accent, fontSize:18 }}>→</span>
              </div>
            </div>
          );
        })
      )}
    </div>
  );

  const DecksTab = () => (
    <div>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:20 }}>
        <div style={{ fontSize:15, fontWeight:700 }}>Finished Decks</div>
        <button style={btn('primary')} onClick={()=>deckFileRef.current?.click()}>+ Upload Deck</button>
      </div>
      <input ref={deckFileRef} type="file" accept=".pptx,.pdf,.key" style={{ display:'none' }} onChange={async(e)=>{ const f=e.target.files[0]; e.target.value=''; if(f) await uploadDeckFile(f,null); }} />
      {clientDecks.length===0?(
        <div style={{ textAlign:'center', padding:'40px', color:C.faint }}>
          <div style={{ fontSize:32, marginBottom:12 }}>🎯</div>
          <div style={{ fontSize:14, fontWeight:600, marginBottom:8 }}>No decks archived yet</div>
          <div style={{ fontSize:13 }}>Upload finished decks here to build your client's history and help Claude build better decks over time.</div>
        </div>
      ):(
        clientDecks.map(d=>{
          const dtype = DECK_TYPES.find(t=>t.id===d.deck_type);
          return (
            <div key={d.id} style={{ ...st.card, marginBottom:10, padding:16, display:'flex', justifyContent:'space-between', alignItems:'center' }}>
              <div>
                <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:4 }}>
                  <span>{dtype?.icon||'📋'}</span>
                  <div style={{ fontSize:14, fontWeight:600 }}>{d.title}</div>
                  <span style={{ fontSize:11, padding:'2px 8px', borderRadius:20, background:'#374151', color:C.muted }}>{dtype?.label||d.deck_type}</span>
                </div>
                <div style={{ fontSize:12, color:C.faint }}>{new Date(d.created_at).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})}</div>
              </div>
              <button style={{ ...btn('sm'), background:'transparent', border:`1px solid ${C.border}`, color:C.muted }} onClick={()=>downloadDeck(d)}>⬇ Download</button>
            </div>
          );
        })
      )}
    </div>
  );

  const NotesTab = () => {
    const [note,setNote]=useState('');
    const [noteType,setNoteType]=useState('general');
    const [saving,setSaving]=useState(false);
    const saveNote=async()=>{ if(!note.trim()) return; setSaving(true); const{data}=await supabase.from('context_notes').insert([{ client_id:selId, note_type:noteType, content:note }]).select().single(); if(data) setClientNotes(prev=>[data,...prev]); setNote(''); setSaving(false); };
    return (
      <div>
        <div style={{ fontSize:15, fontWeight:700, marginBottom:16 }}>Context Notes</div>
        <div style={{ ...st.card, marginBottom:20, background:C.surface }}>
          <div style={{ display:'flex', gap:10, marginBottom:10 }}>
            <select style={{ ...st.input, width:'auto' }} value={noteType} onChange={e=>setNoteType(e.target.value)}>
              {['general','goal_update','competitive','category','meeting'].map(t=><option key={t} value={t}>{t.replace('_',' ')}</option>)}
            </select>
          </div>
          <textarea style={{ ...st.textarea, minHeight:80, marginBottom:10 }} placeholder="Add context, goal updates, competitive intel, meeting notes…" value={note} onChange={e=>setNote(e.target.value)} />
          <button style={{ ...btn('primary'), opacity:(!note.trim()||saving)?0.5:1 }} onClick={saveNote} disabled={!note.trim()||saving}>{saving?'Saving…':'Add Note'}</button>
        </div>
        {clientNotes.map(n=>(
          <div key={n.id} style={{ marginBottom:10, padding:'12px 16px', background:C.surface, borderRadius:10, borderLeft:`3px solid ${C.accent}` }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:6 }}>
              <span style={{ fontSize:11, padding:'2px 8px', borderRadius:20, background:'#374151', color:C.muted }}>{(n.note_type||'general').replace('_',' ')}</span>
              <span style={{ fontSize:11, color:C.faint }}>{new Date(n.created_at).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})}</span>
            </div>
            <div style={{ fontSize:13, color:C.text, lineHeight:1.6, whiteSpace:'pre-wrap' }}>{n.content}</div>
          </div>
        ))}
      </div>
    );
  };

  const NewRequest = () => {
    const [form,setForm]=useState({ deck_type:'sales', title:'', special_instructions:'', due_date:'', requested_by:'' });
    const [selReports,setSelReports]=useState([]);
    const [selTranscripts,setSelTranscripts]=useState([]);
    const [saving,setSaving]=useState(false);
    // ── FIX: use ref instead of state to avoid stale closure on file input onChange ──
    const localUploadTypeRef = useRef(null);
    const localReportRef = useRef(null);

    const relevant = reportsForDeckType(form.deck_type);
    const catGroups = relevant.reduce((acc,r)=>{ (acc[r.cat]=acc[r.cat]||[]).push(r); return acc; },{});

    const handleLocalReport = async (e) => {
      const file = e.target.files[0]; e.target.value = '';
      if (!file || !localUploadTypeRef.current) return;
      const saved = await uploadReportFile(file, localUploadTypeRef.current);
      if (saved) setSelReports(prev => [...prev.filter(id => id !== saved.id), saved.id]);
      localUploadTypeRef.current = null;
    };

    const save = async () => {
      setSaving(true);
      const dtype = DECK_TYPES.find(d => d.id === form.deck_type);
      const title = form.title || `${dtype?.label} — ${client?.name} — ${new Date().toLocaleDateString('en-US', { month:'short', year:'numeric' })}`;

      const { data } = await supabase.from('deck_requests').insert([{
        ...form, title, client_id: selId,
        report_ids: selReports, transcript_ids: selTranscripts, status: 'pending'
      }]).select().single();

      if (data) {
        setClientRequests(prev => [data, ...prev]);
        setSelectedRequest(data);

        // Fire Slack notification (non-blocking)
        try {
          const attachedReportLabels = clientReports
            .filter(r => selReports.includes(r.id))
            .map(r => r.report_label);

          await fetch('/api/slack-notify', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              requestId: data.id,
              clientName: client?.name,
              clientType: client?.type,
              deckTypeLabel: dtype?.label,
              deckTypeIcon: dtype?.icon,
              requestedBy: form.requested_by,
              dueDate: form.due_date,
              specialInstructions: form.special_instructions,
              reportLabels: attachedReportLabels,
              transcriptCount: selTranscripts.length,
            })
          });
        } catch (err) {
          console.error('Slack notify failed (non-fatal):', err);
        }

        setActiveTab('requests');
        setScreen('request-detail');
      }
      setSaving(false);
    };

    return (
      <div style={{ maxWidth:760 }}>
        <button style={{ ...btn('outline'), marginBottom:18 }} onClick={()=>{ setScreen('profile'); setActiveTab('requests'); }}>← Back to {client?.name}</button>
        <div style={{ fontSize:20, fontWeight:800, marginBottom:20 }}>New Deck Request</div>
        <input ref={localReportRef} type="file" accept=".csv,.xlsx,.xls,.txt,.tsv" style={{ display:'none' }} onChange={handleLocalReport} />

        {/* Step 1: Deck Type */}
        <div style={{ ...st.card, marginBottom:16 }}>
          <div style={{ fontSize:14, fontWeight:700, marginBottom:16 }}>1. Deck Type</div>
          <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(180px,1fr))', gap:10 }}>
            {DECK_TYPES.map(d=>(
              <div key={d.id} style={{ padding:'14px', borderRadius:10, border:`2px solid ${form.deck_type===d.id?C.accent:C.border}`, cursor:'pointer', background:form.deck_type===d.id?C.surface:'transparent' }} onClick={()=>setForm(p=>({...p,deck_type:d.id}))}>
                <div style={{ fontSize:20, marginBottom:6 }}>{d.icon}</div>
                <div style={{ fontSize:13, fontWeight:700, color:form.deck_type===d.id?C.accent:C.text }}>{d.label}</div>
                <div style={{ fontSize:11, color:C.faint, marginTop:3 }}>{d.desc}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Step 2: Reports */}
        <div style={{ ...st.card, marginBottom:16 }}>
          <div style={{ fontSize:14, fontWeight:700, marginBottom:6 }}>2. Reports</div>
          <div style={{ fontSize:13, color:C.muted, marginBottom:16 }}>Select reports to include. Upload new ones or use existing files from this client's folder. All optional.</div>
          {Object.entries(catGroups).map(([cat,rpts])=>(
            <div key={cat} style={{ marginBottom:16 }}>
              <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:8 }}>
                <span style={{ width:8, height:8, borderRadius:'50%', background:CAT_COLORS[cat]||C.muted, display:'inline-block' }}></span>
                <span style={{ fontSize:12, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'0.5px' }}>{cat}</span>
              </div>
              <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(250px,1fr))', gap:8 }}>
                {rpts.map(rt=>{
                  const existing = clientReports.filter(r=>r.report_type===rt.id);
                  const anySelected = existing.some(r=>selReports.includes(r.id));
                  return (
                    <div key={rt.id} style={{ background:C.surface, border:`1px solid ${anySelected?C.success:C.border}`, borderRadius:8, padding:'10px 12px' }}>
                      <div style={{ fontSize:12, fontWeight:600, color:C.text, marginBottom:6 }}>{rt.label}</div>
                      {existing.length>0?(
                        existing.map(r=>(
                          <div key={r.id} style={{ display:'flex', alignItems:'center', gap:8, marginBottom:4 }}>
                            <input type="checkbox" checked={selReports.includes(r.id)} onChange={e=>setSelReports(prev=>e.target.checked?[...prev,r.id]:prev.filter(id=>id!==r.id))} style={{ accentColor:C.accent }} />
                            <span style={{ fontSize:11, color:C.faint, flex:1, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.file_name}</span>
                          </div>
                        ))
                      ):(
                        <button style={{ ...btn('sm'), width:'100%', background:'transparent', border:`1px solid ${C.border}`, color:C.muted, fontSize:11 }} onClick={()=>{ localUploadTypeRef.current = rt.id; localReportRef.current?.click(); }}>⬆ Upload</button>
                      )}
                      {existing.length>0&&<button style={{ ...btn('sm'), fontSize:10, marginTop:4, background:'transparent', border:`1px solid ${C.border}`, color:C.faint }} onClick={()=>{ localUploadTypeRef.current = rt.id; localReportRef.current?.click(); }}>+ New version</button>}
                    </div>
                  );
                })}
              </div>
            </div>
          ))}
          {selReports.length>0&&<div style={{ marginTop:12, padding:'8px 12px', background:'#0F2A1D', border:'1px solid #059669', borderRadius:8, fontSize:12, color:C.success }}>✓ {selReports.length} report{selReports.length!==1?'s':''} selected</div>}
        </div>

        {/* Step 3: Transcripts */}
        {clientTranscripts.length>0&&(
          <div style={{ ...st.card, marginBottom:16 }}>
            <div style={{ fontSize:14, fontWeight:700, marginBottom:12 }}>3. Transcripts</div>
            {clientTranscripts.map(t=>(
              <div key={t.id} style={{ display:'flex', alignItems:'center', gap:10, padding:'8px 0', borderBottom:`1px solid ${C.border}` }}>
                <input type="checkbox" checked={selTranscripts.includes(t.id)} onChange={e=>setSelTranscripts(prev=>e.target.checked?[...prev,t.id]:prev.filter(id=>id!==t.id))} style={{ accentColor:C.accent }} />
                <div>
                  <div style={{ fontSize:13, fontWeight:600 }}>{t.title}</div>
                  <div style={{ fontSize:11, color:C.faint }}>{(t.transcript_type||'discovery').replace('_',' ')} · {t.call_date&&new Date(t.call_date).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})}</div>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Step 4: Details */}
        <div style={{ ...st.card, marginBottom:20 }}>
          <div style={{ fontSize:14, fontWeight:700, marginBottom:16 }}>{clientTranscripts.length>0?'4':'3'}. Request Details</div>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14 }}>
            <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Request Title (optional)</label><input style={st.input} value={form.title} onChange={e=>setForm(p=>({...p,title:e.target.value}))} placeholder="Auto-generated if left blank" /></div>
            <div><label style={st.label}>Requested By</label><input style={st.input} value={form.requested_by} onChange={e=>setForm(p=>({...p,requested_by:e.target.value}))} placeholder="Your name" /></div>
            <div><label style={st.label}>Due Date</label><input type="date" style={st.input} value={form.due_date} onChange={e=>setForm(p=>({...p,due_date:e.target.value}))} /></div>
            <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Special Instructions</label><textarea style={{ ...st.textarea, minHeight:80 }} placeholder="Focus on DSP opportunity, compare to Competitor X, client is price-sensitive…" value={form.special_instructions} onChange={e=>setForm(p=>({...p,special_instructions:e.target.value}))} /></div>
          </div>
        </div>

        <div style={{ display:'flex', gap:10 }}>
          <button style={{ ...btn('primary'), opacity:saving?0.5:1 }} onClick={save} disabled={saving}>{saving?'Creating…':'Create Deck Request'}</button>
          <button style={btn('ghost')} onClick={()=>{ setScreen('profile'); setActiveTab('requests'); }}>Cancel</button>
        </div>
      </div>
    );
  };

  const RequestDetail = () => {
    if (!selectedRequest) return null;
    const req = selectedRequest;
    const dtype = DECK_TYPES.find(d=>d.id===req.deck_type);
    const attachedReports = clientReports.filter(r=>(req.report_ids||[]).includes(r.id));
    const attachedTranscripts = clientTranscripts.filter(t=>(req.transcript_ids||[]).includes(t.id));
    const detailDeckRef = useRef(null);

    const updateStatus = async (status) => {
      await supabase.from('deck_requests').update({ status }).eq('id',req.id);
      const updated = { ...req, status };
      setSelectedRequest(updated);
      setClientRequests(prev=>prev.map(r=>r.id===req.id?updated:r));
    };

    return (
      <div style={{ maxWidth:720 }}>
        <button style={{ ...btn('outline'), marginBottom:18 }} onClick={()=>{ setScreen('profile'); setActiveTab('requests'); }}>← Back to {client?.name}</button>
        <input ref={detailDeckRef} type="file" accept=".pptx,.pdf,.key" style={{ display:'none' }} onChange={async(e)=>{ const f=e.target.files[0]; e.target.value=''; if(f) await uploadDeckFile(f,req.id); }} />

        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:20 }}>
          <div>
            <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:6 }}>
              <span style={{ fontSize:22 }}>{dtype?.icon||'📋'}</span>
              <div style={{ fontSize:20, fontWeight:800 }}>{req.title||`${dtype?.label} — ${client?.name}`}</div>
            </div>
            <div style={{ display:'flex', gap:8, alignItems:'center' }}>
              <span style={statusBadge(req.status)}>{(req.status||'pending').replace('_',' ')}</span>
              {req.due_date&&<span style={{ fontSize:12, color:C.faint }}>Due {new Date(req.due_date).toLocaleDateString('en-US',{month:'short',day:'numeric'})}</span>}
              {req.requested_by&&<span style={{ fontSize:12, color:C.faint }}>· {req.requested_by}</span>}
            </div>
          </div>
          <div style={{ display:'flex', gap:8 }}>
            {req.status==='pending'&&<button style={btn('ghost')} onClick={()=>updateStatus('in_progress')}>Mark In Progress</button>}
            {req.status==='in_progress'&&<button style={btn('ghost')} onClick={()=>updateStatus('pending')}>Back to Pending</button>}
          </div>
        </div>

        <div style={{ ...st.card, borderTop:`3px solid ${C.accent}`, marginBottom:16 }}>
          <div style={{ fontSize:15, fontWeight:700, marginBottom:6 }}>🚀 Prepare for Claude</div>
          <div style={{ fontSize:13, color:C.muted, lineHeight:1.6, marginBottom:16 }}>
            Copies a complete briefing to your clipboard — client context, goals, transcripts, and report list. Open your Claude Project, paste this briefing, then upload the report files below. Claude will read every report in full and build the deck.
          </div>
          <button style={{ ...btn('primary'), opacity:briefingCopied?0.5:1 }} onClick={()=>generateBriefing(req)}>
            {briefingCopied?'✓ Copied to Clipboard!':'📋 Copy Briefing for Claude'}
          </button>
          {briefingCopied&&<div style={{ marginTop:12, padding:'10px 14px', background:'#0F2A1D', border:'1px solid #059669', borderRadius:8, fontSize:12, color:C.success }}>
            Briefing copied! Now: 1) Open your Claude Project 2) Paste the briefing 3) Upload the report files below 4) Claude builds the deck
          </div>}
        </div>

        {attachedReports.length>0&&(
          <div style={{ ...st.card, marginBottom:16 }}>
            <div style={{ fontSize:14, fontWeight:700, marginBottom:12 }}>📁 Report Files to Upload to Claude</div>
            <div style={{ fontSize:12, color:C.muted, marginBottom:12 }}>Download each file and upload them all into your Claude Project conversation.</div>
            <div style={{ display:'flex', flexDirection:'column', gap:8 }}>
              {attachedReports.map(r=>(
                <div key={r.id} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'10px 14px', background:C.surface, borderRadius:8 }}>
                  <div>
                    <div style={{ fontSize:13, fontWeight:600 }}>{r.report_label}</div>
                    <div style={{ fontSize:11, color:C.faint }}>{r.file_name}</div>
                  </div>
                  <button style={{ ...btn('sm'), background:'transparent', border:`1px solid ${C.border}`, color:C.muted }} onClick={()=>downloadReport(r)}>⬇ Download</button>
                </div>
              ))}
            </div>
          </div>
        )}

        {attachedTranscripts.length>0&&(
          <div style={{ ...st.card, marginBottom:16 }}>
            <div style={{ fontSize:14, fontWeight:700, marginBottom:12 }}>📝 Transcripts Included</div>
            {attachedTranscripts.map(t=>(
              <div key={t.id} style={{ padding:'8px 0', borderBottom:`1px solid ${C.border}`, fontSize:13 }}>
                <span style={{ fontWeight:600 }}>{t.title}</span>
                <span style={{ color:C.faint, marginLeft:8 }}>{(t.transcript_type||'discovery').replace('_',' ')}</span>
              </div>
            ))}
            <div style={{ fontSize:12, color:C.muted, marginTop:8 }}>Transcript content is included in the Claude briefing above.</div>
          </div>
        )}

        {req.special_instructions&&(
          <div style={{ ...st.card, marginBottom:16 }}>
            <div style={{ fontSize:14, fontWeight:700, marginBottom:8 }}>Special Instructions</div>
            <div style={{ fontSize:13, color:C.text, lineHeight:1.6 }}>{req.special_instructions}</div>
          </div>
        )}

        <div style={{ ...st.card, borderTop:`3px solid ${C.success}` }}>
          <div style={{ fontSize:14, fontWeight:700, marginBottom:6 }}>✅ Upload Finished Deck</div>
          <div style={{ fontSize:13, color:C.muted, marginBottom:14 }}>When the deck is done, upload it here. It'll be saved to {client?.name}'s permanent file and this request will be marked complete.</div>
          <button style={{ ...btn('primary'), background:C.success, opacity:uploadingDeck?0.5:1 }} onClick={()=>detailDeckRef.current?.click()} disabled={uploadingDeck}>
            {uploadingDeck?'⏳ Uploading…':'⬆ Upload Finished Deck'}
          </button>
        </div>
      </div>
    );
  };

  const StyleLibrary = () => (
    <div>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:20 }}>
        <div style={{ fontSize:13, color:C.muted }}>{styleRefs.length} deck{styleRefs.length!==1?'s':''} · {styleRefs.filter(r=>r.active).length} active</div>
        <button style={{ ...btn('primary'), opacity:processingStyleRef?0.5:1 }} onClick={()=>styleRefFileRef.current?.click()} disabled={processingStyleRef}>
          {processingStyleRef?'⏳ Processing…':'+ Add Reference Deck'}
        </button>
      </div>
      <input ref={styleRefFileRef} type="file" accept=".pdf" style={{ display:'none' }} onChange={async(e)=>{ const f=e.target.files[0]; if(f) await processStyleRef(f); e.target.value=''; }} />
      <div style={{ ...st.card, borderLeft:`3px solid ${C.blue}`, marginBottom:20 }}>
        <div style={{ fontSize:14, fontWeight:700, marginBottom:6 }}>How this works</div>
        <div style={{ fontSize:13, color:C.muted, lineHeight:1.7 }}>Upload a past BTR deck PDF once. The AI extracts a style guide — narrative structure, data depth, storytelling approach. That guide is stored here and referenced in the Claude Project when building decks. No re-uploading per session.</div>
      </div>
      {loadingStyleRefs?<div style={{ textAlign:'center', padding:'40px', color:C.faint }}>Loading…</div>:styleRefs.length===0?(
        <div style={{ ...st.card, textAlign:'center', padding:'48px 20px', border:`1px dashed ${C.border}` }}>
          <div style={{ fontSize:40, marginBottom:12 }}>📋</div>
          <div style={{ fontSize:16, fontWeight:700, marginBottom:8 }}>No reference decks yet</div>
          <button style={btn('primary')} onClick={()=>styleRefFileRef.current?.click()}>Upload First Reference Deck</button>
        </div>
      ):(
        styleRefs.map(ref=>(
          <div key={ref.id} style={{ ...st.card, border:`1px solid ${ref.active?C.success:C.border}`, opacity:ref.active?1:0.65, marginBottom:12 }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start' }}>
              <div style={{ flex:1, marginRight:16 }}>
                <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:6 }}>
                  <span style={{ fontSize:20 }}>📋</span>
                  <div style={{ fontSize:15, fontWeight:700 }}>{ref.name}</div>
                  <span style={{ fontSize:11, padding:'2px 8px', borderRadius:20, background:ref.active?'#05966920':'#37415150', color:ref.active?C.success:C.faint, fontWeight:600 }}>{ref.active?'Active':'Inactive'}</span>
                </div>
                <div style={{ fontSize:12, color:C.faint, marginBottom:10 }}>{ref.file_name} · Added {new Date(ref.created_at).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})}</div>
                {ref.style_summary&&<div style={{ fontSize:12, color:C.muted, background:C.surface, borderRadius:8, padding:'10px 14px', lineHeight:1.7, maxHeight:90, overflow:'hidden' }}>{ref.style_summary.slice(0,300)}…</div>}
              </div>
              <div style={{ display:'flex', flexDirection:'column', gap:8, flexShrink:0 }}>
                <button style={{ ...btn('sm'), background:'transparent', border:`1px solid ${C.border}`, color:C.muted }} onClick={()=>toggleStyleRef(ref)}>{ref.active?'Deactivate':'Activate'}</button>
                <button style={{ ...btn('sm'), background:'transparent', border:'1px solid #EF444440', color:C.error }} onClick={()=>deleteStyleRef(ref.id)}>Delete</button>
              </div>
            </div>
          </div>
        ))
      )}
    </div>
  );

  // ── NAV ──────────────────────────────────────────────────────────────────────

  const navItems = [
    { id:'clients', label:'All Clients', icon:'👥' },
    ...(selId&&client?[{ id:'profile', label:client.name.split(' ')[0], icon:'📁' }]:[]),
    ...(screen==='new-client'?[{ id:'new-client', label:'New Client', icon:'➕' }]:[]),
    ...(screen==='new-request'?[{ id:'new-request', label:'New Request', icon:'📋' }]:[]),
    ...(screen==='request-detail'?[{ id:'request-detail', label:'Request', icon:'🚀' }]:[]),
  ];

  const screenTitles = {
    clients: { title:'Client Manager', sub:'All prospects and active clients' },
    'new-client': { title:'New Client', sub:'Create a client profile' },
    'style-library': { title:'Style Library', sub:`${styleRefs.filter(r=>r.active).length} reference deck${styleRefs.filter(r=>r.active).length!==1?'s':''} active` },
    profile: { title:client?.name||'', sub:`${client?.type==='prospect'?'Prospect':'Active Client'} · ${client?.categories||''}` },
    'new-request': { title:'New Deck Request', sub:client?.name||'' },
    'request-detail': { title:selectedRequest?.title||'Deck Request', sub:client?.name||'' },
  };
  const { title='', sub='' } = screenTitles[screen]||{};

  return (
    <div style={st.app}>
      <div style={st.sidebar}>
        <div style={{ padding:'20px 20px 16px', borderBottom:'1px solid #1F2937' }}>
          <div style={{ fontSize:20, fontWeight:900, color:C.accent, letterSpacing:'-0.5px' }}>BTR</div>
          <div style={{ fontSize:10, color:C.faint, fontWeight:600, letterSpacing:'1.2px', textTransform:'uppercase', marginTop:1 }}>Deck Studio</div>
        </div>
        <div style={{ padding:'12px 10px', flex:1 }}>
          {navItems.map(it=><div key={it.id} style={navItem(screen===it.id)} onClick={()=>setScreen(it.id)}><span>{it.icon}</span><span style={{ overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{it.label}</span></div>)}
          <div style={{ borderTop:`1px solid ${C.border}`, marginTop:8, paddingTop:8 }}>
            <div style={navItem(screen==='style-library')} onClick={()=>setScreen('style-library')}>
              <span>📋</span><span>Style Library</span>
              {styleRefs.filter(r=>r.active).length>0&&<span style={{ marginLeft:'auto', fontSize:11, background:C.accent, color:'#fff', borderRadius:20, padding:'1px 7px', fontWeight:700 }}>{styleRefs.filter(r=>r.active).length}</span>}
            </div>
          </div>
        </div>
        <div style={{ padding:'14px 20px', borderTop:'1px solid #1F2937', fontSize:11, color:'#374151' }}>v6.0 · Data Hub</div>
      </div>
      <div style={st.main}>
        <div style={st.header}>
          <div style={{ fontSize:19, fontWeight:700 }}>{title}</div>
          {sub&&<div style={{ fontSize:13, color:C.faint, marginTop:3 }}>{sub}</div>}
        </div>
        <div style={st.content}>
          {screen==='clients'&&<Clients />}
          {screen==='new-client'&&<NewClient />}
          {screen==='style-library'&&<StyleLibrary />}
          {screen==='profile'&&<Profile />}
          {screen==='new-request'&&<NewRequest />}
          {screen==='request-detail'&&<RequestDetail />}
        </div>
      </div>
    </div>
  );
}
