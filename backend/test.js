const fs = require('fs');
const { generatePptx } = require('./src/pptEngine');

(async () => {
  const template = fs.readFileSync('./templates/master.pptx');
  const mcqs = [
    {
      "question": "$$cos2A=\\frac{3}{5}$$ হলে, $$sinA$$ এর মান কত?",
      "optionA": "$$\\pm\\frac{1}{\\sqrt{10}}$$",
      "optionB": "$$\\pm\\frac{1}{\\sqrt{5}}$$",
      "optionC": "$$\\pm\\frac{3}{5}$$",
      "optionD": "$$\\frac{2}{5}$$",
      "correct": "B"
    },
    {
      "question": "$$tan\\beta=\\frac{q}{p}$$ হলে, $$cos2\\beta$$ এর মান কত?",
      "optionA": "$$\\frac{2qp}{p^2-q^2}$$",
      "optionB": "$$\\frac{2qp}{p^2+q^2}$$",
      "optionC": "$$\\frac{p^2-q^2}{p^2+q^2}$$",
      "optionD": "$$\\frac{p^2+q^2}{p^2-q^2}$$",
      "correct": "C"
    },
    {
      "question": "যদি $$cos\\theta=\\frac{1}{2}\\left(x+\\frac{1}{x}\\right)$$ হয়, তবে $$cos2\\theta=$$ ?",
      "optionA": "$$-\\frac{1}{2}\\left(x+\\frac{1}{x}\\right)^2$$",
      "optionB": "$$\\frac{1}{2}\\left(x-\\frac{1}{x}\\right)^2$$",
      "optionC": "$$\\frac{1}{2}\\left(\\frac{1}{x^2}-x\\right)$$",
      "optionD": "$$\\frac{1}{2}\\left(x^2+\\frac{1}{x^2}\\right)$$",
      "correct": "D"
    },
    {
      "question": "$$cos^4\\theta-sin^4\\theta=$$ ?",
      "optionA": "$$2cos^2\\theta-1$$",
      "optionB": "$$2sin^2\\theta-1$$",
      "optionC": "$$2tan^2\\theta-1$$",
      "optionD": "$$2sec^2\\theta-1$$",
      "correct": "A"
    },
    {
      "question": "$$tan20^{\\circ} tan40^{\\circ} tan80^{\\circ}=$$ ?",
      "optionA": "$$\\sqrt{3}$$",
      "optionB": "$$-\\sqrt{3}$$",
      "optionC": "$$\\frac{1}{\\sqrt{3}}$$",
      "optionD": "$$-\\frac{1}{\\sqrt{3}}$$",
      "correct": "A"
    },
    {
      "question": "\\frac{1+cos2\\theta}{sin2\\theta}=$$ ?",
      "optionA": "$$cot\\theta$$",
      "optionB": "$$tan\\theta$$",
      "optionC": "$$sec\\theta$$",
      "optionD": "$$cosec\\theta$$",
      "correct": "A"
    },
    {
      "question": "$$k, l$$ এর কোন মানের জন্য $$5sin(k\\theta)=(10l+9)sin\\theta+(15l+6)cos\\theta$$ একটি অভেদ হবে?",
      "optionA": "$$1, -\\frac{2}{5}$$",
      "optionB": "$$-1, -\\frac{2}{5}$$",
      "optionC": "$$-1, \\frac{2}{5}$$",
      "optionD": "$1, -\\frac{5}{2}$$",
      "correct": "A"
    },
    {
      "question": "$$\\frac{sinA}{1+cosA}+\\frac{1-cosAsinA}{2}=$$ ?",
      "optionA": "$$2cosecA$$",
      "optionB": "$$2tan\\frac{A}{2}$$",
      "optionC": "$$2cotA$$",
      "optionD": "$$2secA$$",
      "correct": "B"
    },
    {
      "question": "$$tan\\theta=\\frac{1}{2}$$ হলে, $$10sin2\\theta-6tan2\\theta+5cos2\\theta=$$ ?",
      "optionA": "$$12$$",
      "optionB": "$$1$$",
      "optionC": "$$3$$",
      "optionD": "$$-13$$",
      "correct": "C"
    },
    {
      "question": "$$\\frac{\\sqrt{3}}{sin20^{\\circ}}-\\frac{1}{cos20^{\\circ}}=$$ ?",
      "optionA": "$$0$$",
      "optionB": "$$2$$",
      "optionC": "$$-\\frac{1}{2}$$",
      "optionD": "$$4$$",
      "correct": "D"
    }
  ];
  try {
    const out = await generatePptx(template, mcqs);
    fs.writeFileSync('./test_out.pptx', out);
    console.log('Successfully wrote', out.length, 'bytes to test_out.pptx');
  } catch (err) {
    console.error('Error generating pptx:', err);
  }
})();
