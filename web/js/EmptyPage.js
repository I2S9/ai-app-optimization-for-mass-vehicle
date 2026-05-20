export default {
  name: 'EmptyPage',
  props: {
    title: { type: String, required: true },
  },
  template: `
    <div class="empty-page">
      <h2>{{ title }}</h2>
      <p>This module will be implemented in a later phase.</p>
    </div>
  `,
};
