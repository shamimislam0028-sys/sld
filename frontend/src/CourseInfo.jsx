import React, { useRef } from 'react';

export default function CourseInfo({ mnemonicFile, setMnemonicFile, iconFile, setIconFile, onNext }) {
  const iconRef = useRef(null);
  const mnemRef = useRef(null);
  const iconUrl = iconFile ? URL.createObjectURL(iconFile) : null;
  const mnemUrl = mnemonicFile ? URL.createObjectURL(mnemonicFile) : null;

  const canProceed = mnemonicFile && iconFile;

  function pick(setter) {
    return (e) => {
      const f = e.target.files?.[0];
      if (!f) return;
      const ok = ['image/png', 'image/jpeg', 'image/svg+xml'].includes(f.type);
      if (!ok) {
        alert('Please upload a PNG, JPG, or SVG.');
        return;
      }
      setter(f);
    };
  }

  return (
    <div className="card">
      <h2>Course Information</h2>
      <p className="muted">These appear on every generated slide.</p>

      <label className="field">
        <span>Course Mnemonic image (PNG / JPG / SVG) — fills the top-right box</span>
        <div className="icon-upload">
          <button type="button" className="btn ghost" onClick={() => mnemRef.current?.click()}>
            {mnemonicFile ? 'Change image' : 'Choose file'}
          </button>
          <input ref={mnemRef} type="file" accept=".png,.jpg,.jpeg,.svg" onChange={pick(setMnemonicFile)} hidden />
          {mnemUrl && <img className="icon-preview" src={mnemUrl} alt="mnemonic preview" />}
          {mnemonicFile && <span className="muted small">{mnemonicFile.name}</span>}
        </div>
      </label>

      <label className="field">
        <span>Course Icon (PNG / JPG / SVG) — fills the top-left box</span>
        <div className="icon-upload">
          <button type="button" className="btn ghost" onClick={() => iconRef.current?.click()}>
            {iconFile ? 'Change icon' : 'Choose file'}
          </button>
          <input ref={iconRef} type="file" accept=".png,.jpg,.jpeg,.svg" onChange={pick(setIconFile)} hidden />
          {iconUrl && <img className="icon-preview" src={iconUrl} alt="icon preview" />}
          {iconFile && <span className="muted small">{iconFile.name}</span>}
        </div>
      </label>

      <div className="actions end">
        <button className="btn primary" disabled={!canProceed} onClick={onNext}>
          Continue to questions →
        </button>
      </div>
    </div>
  );
}
