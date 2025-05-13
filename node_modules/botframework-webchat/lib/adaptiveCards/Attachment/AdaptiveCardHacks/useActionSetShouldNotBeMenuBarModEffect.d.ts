import type { AdaptiveCard } from 'adaptivecards';
/**
 * Accessibility: ActionSet should not be menu bar.
 *
 * Menu bar is not accessible through screen reader keyboard shortcut keys:
 *
 * - <kbd>B</kbd> will jump to next button, which the end-user can quickly the chosen action;
 * - <kbd>F</kbd> will jump to next form field, which is very similar but also jump to text fields;
 * - There are no keyboard shortcut keys for menu.
 *
 * Marking action button as menu item in a menu bar hurts accessibility. End-user will not be able to jump to buttons quickly.
 *
 * Thus, ActionSet should not be menu bar.
 */
export default function useActionShouldBePushButtonModEffect(adaptiveCard: AdaptiveCard): readonly [(cardElement: HTMLElement) => void, () => void];
//# sourceMappingURL=useActionSetShouldNotBeMenuBarModEffect.d.ts.map