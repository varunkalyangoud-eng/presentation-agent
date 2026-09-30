const form = document.getElementById('presentation-form');
const statusBox = document.getElementById('status');
const cardsContainer = document.getElementById('cards');
const slidesContainer = document.getElementById('slides');
const downloadArea = document.getElementById('download-area');

function renderCards(cards) {
  if (!cards || !cards.length) {
    cardsContainer.innerHTML = '';
    return;
  }

  cardsContainer.innerHTML = cards.map((card) => `
    <div class="card">
      <div class="label">${card.label}</div>
      <div class="value">${card.value}</div>
    </div>
  `).join('');
}

function renderSlides(slides) {
  if (!slides || !slides.length) {
    slidesContainer.innerHTML = '';
    return;
  }

  slidesContainer.innerHTML = slides.map((slide) => `
    <article class="slide">
      <h3>${slide.title}</h3>
      <ul>
        ${slide.bullets.map((bullet) => `<li>${bullet}</li>`).join('')}
      </ul>
    </article>
  `).join('');
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  const formData = new FormData(form);
  const fileInput = document.getElementById('files');

  if (!fileInput.files.length) {
    statusBox.className = 'status neutral';
    statusBox.textContent = 'Please select at least one file to generate a presentation.';
    return;
  }

  statusBox.className = 'status neutral';
  statusBox.textContent = 'Generating deck...';

  try {
    const response = await fetch('/api/generate', {
      method: 'POST',
      body: formData
    });

    const data = await response.json();
    if (!response.ok || data.error) {
      throw new Error(data.error || 'Generation failed.');
    }

    renderCards(data.cards);
    renderSlides(data.slides);

    statusBox.className = 'status success';
    statusBox.textContent = 'Presentation generated successfully.';

    downloadArea.innerHTML = `
      <a class="download-link" href="${data.pptx_url}" target="_blank" rel="noreferrer">
        Download PowerPoint
      </a>
    `;
  } catch (error) {
    statusBox.className = 'status neutral';
    statusBox.textContent = error.message || 'Something went wrong while generating the deck.';
  }
});
