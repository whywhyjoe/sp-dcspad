// Font-backed <fluent-icon> runtime for preview documents.
// Classic, dependency-free script: the runner injects it after the configured
// Fluent font CSS files and before user JavaScript.

(function () {
  if (customElements.get('fluent-icon')) return;

  function iconClass(name) {
    var match = String(name || '').match(
      /^([a-z0-9-]+)-(\d+)-(regular|filled|light)(?:-(ltr|rtl))?$/i
    );
    if (!match) return null;
    return {
      className: 'icon-ic_fluent_'
        + match[1].toLowerCase().replace(/-/g, '_')
        + '_' + match[2] + '_' + match[3].toLowerCase()
        + (match[4] ? '_' + match[4].toLowerCase() : ''),
      sourceSize: Number(match[2]),
    };
  }

  class FluentIcon extends HTMLElement {
    static get observedAttributes() { return ['name', 'label', 'class']; }

    connectedCallback() { this.render(); }
    attributeChangedCallback() { if (this.isConnected) this.render(); }

    render() {
      var resolved = iconClass(this.getAttribute('name'));
      this.replaceChildren();
      if (!resolved) return;

      var glyph = document.createElement('i');
      glyph.className = resolved.className;
      glyph.setAttribute('aria-hidden', 'true');
      glyph.style.display = 'inline-block';
      glyph.style.fontSize = resolved.sourceSize + 'px';
      glyph.style.lineHeight = '1';

      var sizeClass = Array.prototype.find.call(
        this.classList,
        function (name) { return /^icon--\d+$/.test(name); }
      );
      if (sizeClass) glyph.style.fontSize = sizeClass.slice(6) + 'px';

      var label = this.getAttribute('label');
      if (label) {
        this.setAttribute('role', 'img');
        this.setAttribute('aria-label', label);
        this.removeAttribute('aria-hidden');
      } else if (!this.hasAttribute('aria-label')) {
        this.setAttribute('aria-hidden', 'true');
        if (this.getAttribute('role') === 'img') this.removeAttribute('role');
      }

      this.appendChild(glyph);
    }
  }

  customElements.define('fluent-icon', FluentIcon);
})();
