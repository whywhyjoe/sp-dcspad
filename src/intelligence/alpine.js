// Alpine.js v3 intelligence shared by Monaco's JavaScript and HTML surfaces.
// This intentionally describes the stable public/core API. Plugin-only
// directives and magics belong to separate packs tied to their runtimes.

export const ALPINE_PACK_ID = 'alpine-3';

export const ALPINE_JS_LIBRARIES = [{
  filePath: 'file:///node_modules/@types/dcspad-alpine/index.d.ts',
  content: `/**
 * Alpine.js v3 public browser API.
 * @see https://alpinejs.dev/globals/
 */
interface AlpineMagicProperties {
  /** The current DOM element. */
  readonly $el: HTMLElement;
  /** Elements marked with x-ref in the current component. */
  readonly $refs: Record<string, HTMLElement>;
  /** Global stores registered with Alpine.store(). */
  readonly $store: Record<string, any>;
  /** The nearest Alpine component root. */
  readonly $root: HTMLElement;
  /** The current merged Alpine data scope. */
  readonly $data: Record<string, any>;
  /** Watch a dot-notated component property for changes. */
  $watch<T = any>(property: string, callback: (value: T, oldValue: T) => void): () => void;
  /** Dispatch a bubbling CustomEvent from the current element. */
  $dispatch(name: string, detail?: any): boolean;
  /** Run work after Alpine has flushed reactive DOM updates. */
  $nextTick(callback?: () => void): Promise<void>;
  /** Generate a component-scoped element id. */
  $id(name: string, key?: string | number): string;
}

type AlpinePlugin = (alpine: AlpineStatic) => void;
type AlpineDataProvider<T extends object = Record<string, any>> =
  (...parameters: any[]) => T & ThisType<T & AlpineMagicProperties>;
type AlpineBindProvider =
  () => Record<string, any> & ThisType<Record<string, any> & AlpineMagicProperties>;

interface AlpineStatic {
  /** The loaded Alpine runtime version. */
  readonly version: string;
  /** Start Alpine and initialize the current document. */
  start(): void;

  /**
   * Register a reusable x-data provider.
   * @see https://alpinejs.dev/globals/alpine-data
   */
  data<T extends object>(name: string, provider: AlpineDataProvider<T>): void;

  /**
   * Read a global store.
   * @see https://alpinejs.dev/globals/alpine-store
   */
  store<T = any>(name: string): T;
  /** Register or replace a global store. */
  store<T extends object>(name: string, value: T & ThisType<T & AlpineMagicProperties>): T;

  /**
   * Register a reusable x-bind object.
   * @see https://alpinejs.dev/globals/alpine-bind
   */
  bind(name: string, provider: AlpineBindProvider): void;

  /** Install one or more Alpine plugins before Alpine.start(). */
  plugin(plugin: AlpinePlugin | AlpinePlugin[]): void;
  /** Register a custom magic property. */
  magic(name: string, callback: (element: HTMLElement, utilities: Record<string, any>) => any): void;
  /** Register a custom x-* directive. */
  directive(name: string, callback: (...args: any[]) => void): { before(other: string): void };
  /** Change or return Alpine's directive prefix. */
  prefix(prefix: string): void;
  prefixed(subject?: string): string;

  /** Create a deeply reactive proxy. */
  reactive<T extends object>(target: T): T;
  /** Return the original object behind a reactive proxy. */
  raw<T>(value: T): T;
  /** Run and track a reactive effect. */
  effect(callback: () => void): any;
  /** Stop a reactive effect. */
  release(effect: any): void;
  /** Watch a reactive getter. */
  watch<T>(getter: () => T, callback: (value: T, oldValue: T) => void): () => void;
  /** Schedule work after Alpine's next DOM update. */
  nextTick(callback?: () => void): Promise<void>;

  /** Return the merged Alpine data scope for an element. */
  $data<T extends object = Record<string, any>>(element: Element): T;
  /** Initialize Alpine behavior under an element added dynamically. */
  initTree(element: Element): void;
  /** Tear down Alpine behavior under an element. */
  destroyTree(element: Element): void;
}

declare const Alpine: AlpineStatic;

interface Window {
  Alpine: AlpineStatic;
}
`,
}];

const directive = (name, description, {
  insertText = `${name}="\${1:expression}"`,
  url = `https://alpinejs.dev/directives/${name.slice(2).split(':')[0]}`,
} = {}) => ({ name, description, insertText, url });

export const ALPINE_DIRECTIVES = [
  directive('x-data', 'Declares a new Alpine component and its reactive state.', {
    insertText: `x-data="{ \${1:open}: \${2:false} }"`,
  }),
  directive('x-init', 'Runs an expression while Alpine initializes the element.'),
  directive('x-show', 'Toggles element visibility from a truthy expression.'),
  directive('x-bind', 'Binds an object of attributes, or a reusable Alpine.bind() provider.'),
  directive('x-bind:class', 'Reactively binds the element class attribute.'),
  directive('x-bind:aria-expanded', 'Reactively binds aria-expanded.', {
    insertText: `x-bind:aria-expanded="\${1:open}"`,
    url: 'https://alpinejs.dev/directives/bind',
  }),
  directive('x-bind:aria-selected', 'Reactively binds aria-selected.', {
    insertText: `x-bind:aria-selected="\${1:selected}"`,
    url: 'https://alpinejs.dev/directives/bind',
  }),
  directive('x-bind:aria-pressed', 'Reactively binds aria-pressed.', {
    insertText: `x-bind:aria-pressed="\${1:pressed}"`,
    url: 'https://alpinejs.dev/directives/bind',
  }),
  directive('x-on', 'Attaches an event listener. Add an event name, such as x-on:click.'),
  directive('x-on:click', 'Runs an expression when the element is clicked.'),
  directive('x-on:submit.prevent', 'Prevents form submission and runs an expression.', {
    url: 'https://alpinejs.dev/directives/on',
  }),
  directive('x-on:click.outside', 'Runs when a click occurs outside the element.', {
    url: 'https://alpinejs.dev/directives/on',
  }),
  directive('x-on:keydown.escape.window', 'Runs on Escape keydown from window.', {
    url: 'https://alpinejs.dev/directives/on',
  }),
  directive('x-text', 'Sets textContent from an expression.'),
  directive('x-html', 'Sets innerHTML from an expression. Only use trusted content.'),
  directive('x-model', 'Creates two-way binding between form state and component data.'),
  directive('x-modelable', 'Exposes an internal property to an outer x-model binding.'),
  directive('x-for', 'Repeats a template for each item in an iterable.', {
    insertText: `x-for="(\${1:item}, \${2:index}) in \${3:items}"`,
  }),
  directive('x-transition', 'Adds Alpine transition classes around x-show changes.', {
    insertText: 'x-transition',
  }),
  directive('x-effect', 'Re-runs an expression whenever its reactive dependencies change.'),
  directive('x-ignore', 'Prevents Alpine from initializing this element subtree.', {
    insertText: 'x-ignore',
  }),
  directive('x-ref', 'Names an element for access through $refs.', {
    insertText: `x-ref="\${1:name}"`,
  }),
  directive('x-cloak', 'Keeps an element hidden until Alpine initializes it.', {
    insertText: 'x-cloak',
  }),
  directive('x-teleport', 'Moves a template to another DOM location.', {
    insertText: `x-teleport="\${1:body}"`,
  }),
  directive('x-if', 'Conditionally adds or removes a template from the DOM.'),
  directive('x-id', 'Declares names used by the component-scoped $id() helper.', {
    insertText: `x-id="['\${1:control}']"`,
  }),
];

const shorthand = (name, insertText, description, url) => ({
  name, insertText, description, url,
});

export const ALPINE_SHORTHANDS = [
  shorthand('@click', `@click="\${1:expression}"`, 'Shorthand for x-on:click.', 'https://alpinejs.dev/directives/on'),
  shorthand('@click.outside', `@click.outside="\${1:open = false}"`, 'Runs when clicking outside the element.', 'https://alpinejs.dev/directives/on'),
  shorthand('@submit.prevent', `@submit.prevent="\${1:submit()}"`, 'Prevents submission and runs an expression.', 'https://alpinejs.dev/directives/on'),
  shorthand('@keydown.escape.window', `@keydown.escape.window="\${1:open = false}"`, 'Runs on Escape from window.', 'https://alpinejs.dev/directives/on'),
  shorthand(':class', `:class="{ '\${1:is-active}': \${2:active} }"`, 'Shorthand for x-bind:class.', 'https://alpinejs.dev/directives/bind'),
  shorthand(':disabled', `:disabled="\${1:disabled}"`, 'Shorthand for x-bind:disabled.', 'https://alpinejs.dev/directives/bind'),
  shorthand(':aria-expanded', `:aria-expanded="\${1:open}"`, 'Shorthand for x-bind:aria-expanded.', 'https://alpinejs.dev/directives/bind'),
  shorthand(':aria-selected', `:aria-selected="\${1:selected}"`, 'Shorthand for x-bind:aria-selected.', 'https://alpinejs.dev/directives/bind'),
  shorthand(':aria-pressed', `:aria-pressed="\${1:pressed}"`, 'Shorthand for x-bind:aria-pressed.', 'https://alpinejs.dev/directives/bind'),
];

const magic = (name, description, insertText = name) => ({
  name,
  description,
  insertText,
  url: name === '$event'
    ? 'https://alpinejs.dev/directives/on'
    : `https://alpinejs.dev/magics/${name.slice(1).toLowerCase()}`,
});

export const ALPINE_MAGICS = [
  magic('$el', 'The current DOM element.'),
  magic('$refs', 'Elements marked with x-ref in the current component.', `$refs.\${1:name}`),
  magic('$store', 'Global stores registered with Alpine.store().', `$store.\${1:name}`),
  magic('$watch', 'Watches a component property for changes.', `$watch('\${1:property}', (\${2:value}, \${3:oldValue}) => { \${0} })`),
  magic('$dispatch', 'Dispatches a bubbling CustomEvent from the current element.', `$dispatch('\${1:event}', { \${0} })`),
  magic('$nextTick', 'Runs work after Alpine flushes reactive DOM updates.', `$nextTick(() => { \${0} })`),
  magic('$root', 'The nearest Alpine component root element.'),
  magic('$data', 'The current merged Alpine data scope.'),
  magic('$id', 'Generates a component-scoped element id.', `$id('\${1:name}')`),
  magic('$event', 'The native event available inside x-on expressions.'),
];

const htmlAttribute = (item) => ({
  name: item.name,
  description: item.description,
  references: [{ name: 'Alpine.js documentation', url: item.url }],
});

export const ALPINE_HTML_DATA = {
  version: 1.1,
  globalAttributes: [
    ...ALPINE_DIRECTIVES.map(htmlAttribute),
    ...ALPINE_SHORTHANDS.map(htmlAttribute),
  ],
};

function completionRange(monaco, model, position, prefix) {
  const end = model.getOffsetAt(position);
  const start = Math.max(0, end - prefix.length);
  return new monaco.Range(
    model.getPositionAt(start).lineNumber,
    model.getPositionAt(start).column,
    position.lineNumber,
    position.column,
  );
}

function alpineAttributeValue(fragment) {
  const match = fragment.match(
    /(?:^|\s)(x-[\w:.-]+|@[\w:.-]+|:[\w:.-]+)\s*=\s*(["'])([\s\S]*)$/,
  );
  if (!match || match[3].includes(match[2])) return null;
  return match[3];
}

export function createAlpineHtmlCompletionProvider(monaco, isEnabled) {
  return {
    triggerCharacters: ['x', '-', '@', ':', '$', '.'],
    provideCompletionItems(model, position) {
      if (!isEnabled()) return { suggestions: [] };

      const offset = model.getOffsetAt(position);
      const before = model.getValue().slice(Math.max(0, offset - 6000), offset);
      const lastOpen = before.lastIndexOf('<');
      const lastClose = before.lastIndexOf('>');
      if (lastOpen <= lastClose) return { suggestions: [] };
      const fragment = before.slice(lastOpen + 1);

      const value = alpineAttributeValue(fragment);
      if (value !== null) {
        const magicPrefix = value.match(/\$[A-Za-z]*$/)?.[0];
        if (!magicPrefix) return { suggestions: [] };
        const range = completionRange(monaco, model, position, magicPrefix);
        return {
          suggestions: ALPINE_MAGICS.map((item) => ({
            label: item.name,
            kind: monaco.languages.CompletionItemKind.Variable,
            detail: 'Alpine magic property',
            documentation: {
              value: `${item.description}\n\n[Alpine.js documentation](${item.url})`,
            },
            insertText: item.insertText,
            insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
            range,
          })),
        };
      }

      const prefix = fragment.match(/(?:^|\s)(x-[^\s"'=<>]*|@[^\s"'=<>]*|:[^\s"'=<>]*)$/)?.[1];
      if (!prefix) return { suggestions: [] };
      const source = prefix.startsWith('x-') ? ALPINE_DIRECTIVES : ALPINE_SHORTHANDS;
      const range = completionRange(monaco, model, position, prefix);
      return {
        suggestions: source.map((item) => ({
          label: item.name,
          kind: monaco.languages.CompletionItemKind.Property,
          detail: prefix.startsWith('x-') ? 'Alpine directive' : 'Alpine shorthand',
          documentation: {
            value: `${item.description}\n\n[Alpine.js documentation](${item.url})`,
          },
          insertText: item.insertText,
          filterText: item.name,
          insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
          range,
        })),
      };
    },
  };
}
