import React from 'react';
import { LETTERS, validateMcq } from './utils.js';

export default function QuestionCard({
  mcq,
  index,
  onChange,
  onDuplicate,
  onDelete,
  dragHandlers,
  isDragging,
}) {
  const errs = validateMcq(mcq, index);
  const set = (field) => (e) => onChange({ ...mcq, [field]: e.target.value });

  return (
    <div
      className={`qcard ${isDragging ? 'dragging' : ''}`}
      draggable
      {...dragHandlers}
    >
      <div className="qcard-head">
        <span className="drag" title="Drag to reorder">⋮⋮</span>
        <strong>Question {index + 1}</strong>
        <div className="spacer" />
        <button className="btn tiny" onClick={() => onDuplicate(mcq.id)}>Duplicate</button>
        <button className="btn tiny danger" onClick={() => onDelete(mcq.id)}>Delete</button>
      </div>

      <textarea
        className="p-text"
        rows={2}
        placeholder="Enter passage (উদ্দীপক) if any..."
        value={mcq.passage || ''}
        onChange={set('passage')}
        style={{ marginBottom: '8px', borderLeft: '3px solid var(--magenta)', paddingLeft: '8px' }}
      />

      <textarea
        className="q-text"
        rows={2}
        placeholder="Enter the question…"
        value={mcq.question}
        onChange={set('question')}
      />

      <div className="options-grid">
        {LETTERS.map((L) => (
          <label key={L} className="opt-field">
            <span className="opt-tag">{L}</span>
            <input
              type="text"
              placeholder={`Option ${L}`}
              value={mcq['option' + L]}
              onChange={set('option' + L)}
            />
          </label>
        ))}
      </div>

      <div className="qcard-foot">
        <label className="field inline">
          <span>Correct answer</span>
          <select value={mcq.correct} onChange={set('correct')}>
            {LETTERS.map((L) => (
              <option key={L} value={L}>{L}</option>
            ))}
          </select>
        </label>
        {errs.length > 0 && (
          <span className="inline-errs">{errs.join(' · ')}</span>
        )}
      </div>
    </div>
  );
}
