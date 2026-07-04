import React, { useEffect, useRef, useState } from 'react';
import CourseInfo from './CourseInfo.jsx';
import QuestionCard from './QuestionCard.jsx';
import {
  blankMcq,
  validateAll,
  saveState,
  loadState,
  clearState,
  exportJson,
  parseImportedJson,
} from './utils.js';

const API = '/api/generate';

export default function App() {
  const [step, setStep] = useState(1); // 1 = course info, 2 = questions
  const [mnemonicFile, setMnemonicFile] = useState(null);
  const [iconFile, setIconFile] = useState(null);
  const [mcqs, setMcqs] = useState([blankMcq()]);
  const [busy, setBusy] = useState(false);
  const [errors, setErrors] = useState([]);
  const importRef = useRef(null);
  const docxImportRef = useRef(null);
  const dragIndex = useRef(null);
  const [dragging, setDragging] = useState(null);

  // hydrate from localStorage once
  useEffect(() => {
    const saved = loadState();
    if (saved) {
      if (saved.mcqs?.length) setMcqs(saved.mcqs);
    }
  }, []);

  // autosave (mnemonic + questions; icon can't be serialized)
  useEffect(() => {
    saveState({ mcqs });
  }, [mcqs]);

  /* ----------------------- question operations ---------------------- */
  const updateMcq = (updated) =>
    setMcqs((list) => list.map((m) => (m.id === updated.id ? updated : m)));

  const addMcq = () => setMcqs((list) => [...list, blankMcq()]);

  const duplicateMcq = (id) =>
    setMcqs((list) => {
      const idx = list.findIndex((m) => m.id === id);
      if (idx === -1) return list;
      const copy = { ...list[idx], id: crypto.randomUUID() };
      const next = [...list];
      next.splice(idx + 1, 0, copy);
      return next;
    });

  const deleteMcq = (id) =>
    setMcqs((list) => (list.length === 1 ? list : list.filter((m) => m.id !== id)));

  /* --------------------------- drag reorder ------------------------- */
  const makeDragHandlers = (index) => ({
    onDragStart: () => {
      dragIndex.current = index;
      setDragging(index);
    },
    onDragOver: (e) => e.preventDefault(),
    onDrop: () => {
      const from = dragIndex.current;
      if (from === null || from === index) return;
      setMcqs((list) => {
        const next = [...list];
        const [moved] = next.splice(from, 1);
        next.splice(index, 0, moved);
        return next;
      });
      dragIndex.current = null;
      setDragging(null);
    },
    onDragEnd: () => {
      dragIndex.current = null;
      setDragging(null);
    },
  });

  /* ----------------------------- import ----------------------------- */
  const handleImport = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = parseImportedJson(reader.result);
        if (!parsed.length) throw new Error('No questions found.');
        setMcqs(parsed);
        setErrors([]);
      } catch (err) {
        setErrors(['Import failed: ' + err.message]);
      }
    };
    reader.readAsText(f);
    e.target.value = '';
  };

  const handleDocxImport = async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    e.target.value = '';

    setBusy(true);
    setErrors([]);
    try {
      const formData = new FormData();
      formData.append('docx', f);

      const resp = await fetch('/api/parse-docx', {
        method: 'POST',
        body: formData,
      });

      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error || 'DOCX parsing failed.');

      if (!data.mcqs || !data.mcqs.length) throw new Error('No questions found in the DOCX.');

      // Assign fresh UUIDs to every imported question
      const imported = data.mcqs.map((m) => ({ ...blankMcq(), ...m }));
      setMcqs(imported);
    } catch (err) {
      setErrors(['DOCX Import failed: ' + err.message]);
    } finally {
      setBusy(false);
    }
  };

  /* ---------------------------- generate ---------------------------- */
  const generate = async () => {
    const errs = validateAll({ mnemonicFile, iconFile, mcqs });
    setErrors(errs);
    if (errs.length) return;

    setBusy(true);
    try {
      const fd = new FormData();
      fd.append('mnemonic', mnemonicFile);
      fd.append('icon', iconFile);
      fd.append('mcqs', JSON.stringify(mcqs));

      const res = await fetch(API, { method: 'POST', body: fd });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || `Server error (${res.status})`);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'MCQ_deck.pptx';
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      setErrors([err.message]);
    } finally {
      setBusy(false);
    }
  };

  /* ------------------------------ render ---------------------------- */
  const answered = mcqs.filter(
    (m) =>
      m.question.trim() &&
      m.optionA.trim() &&
      m.optionB.trim() &&
      m.optionC.trim() &&
      m.optionD.trim()
  ).length;
  const progress = mcqs.length ? Math.round((answered / mcqs.length) * 100) : 0;

  return (
    <div className="app">
      <header className="app-head">
        <h1>MCQ PowerPoint Generator</h1>
        <div className="steps">
          <span className={step === 1 ? 'on' : ''}>1 · Course</span>
          <span className="sep">—</span>
          <span className={step === 2 ? 'on' : ''}>2 · Questions</span>
        </div>
      </header>

      {errors.length > 0 && (
        <div className="errors">
          {errors.map((e, i) => (
            <div key={i}>• {e}</div>
          ))}
        </div>
      )}

      {step === 1 && (
        <CourseInfo
          mnemonicFile={mnemonicFile}
          setMnemonicFile={setMnemonicFile}
          iconFile={iconFile}
          setIconFile={setIconFile}
          onNext={() => setStep(2)}
        />
      )}

      {step === 2 && (
        <>
          <div className="toolbar">
            <button className="btn ghost" onClick={() => setStep(1)}>← Course info</button>
            <div className="spacer" />
            <span className="counter">{answered}/{mcqs.length} complete</span>
            <button className="btn ghost" onClick={() => exportJson(mcqs)}>Export JSON</button>
            <button className="btn ghost" onClick={() => importRef.current?.click()}>Import JSON</button>
            <input ref={importRef} type="file" accept=".json" hidden onChange={handleImport} />
            <button className="btn ghost" onClick={() => docxImportRef.current?.click()}>Import DOCX</button>
            <input ref={docxImportRef} type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" hidden onChange={handleDocxImport} />
            <button
              className="btn ghost"
              onClick={() => {
                if (confirm('Clear all questions and saved data?')) {
                  clearState();
                  setMcqs([blankMcq()]);
                  setMnemonicFile(null);
                }
              }}
            >
              Reset
            </button>
          </div>

          <div className="progress">
            <div className="bar" style={{ width: `${progress}%` }} />
          </div>

          <div className="qlist">
            {mcqs.map((m, i) => (
              <QuestionCard
                key={m.id}
                mcq={m}
                index={i}
                onChange={updateMcq}
                onDuplicate={duplicateMcq}
                onDelete={deleteMcq}
                dragHandlers={makeDragHandlers(i)}
                isDragging={dragging === i}
              />
            ))}
          </div>

          <div className="actions">
            <button className="btn" onClick={addMcq}>+ Add question</button>
            <div className="spacer" />
            <button className="btn primary" onClick={generate} disabled={busy}>
              {busy ? 'Generating…' : 'Generate PowerPoint'}
            </button>
          </div>
        </>
      )}

      <footer className="app-foot muted small">
        Slides are generated from your uploaded template, verbatim. Two slides per
        question (question + answer).
      </footer>
    </div>
  );
}
