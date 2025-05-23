/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  EducationalActivity,
  PersonAnnualEvent,
  PersonInterest,
  PhysicalAddress,
  Profile
} from '@microsoft/microsoft-graph-types-beta';
import { html, TemplateResult, nothing } from 'lit';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { styles } from './mgt-profile-css';
import { strings } from './strings';
import { registerComponent } from '@microsoft/mgt-element';
import { property } from 'lit/decorators.js';

export var registerMgtProfileComponent = () => registerComponent('profile', MgtProfile);

/**
 * The user profile subsection of the person card
 *
 * @export
 * @class MgtProfile
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --token-overflow-color - {Color} Color of the text showing more undisplayed values i.e. +3 more
 */
export class MgtProfile extends BasePersonCardSection {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  protected get strings() {
    return strings;
  }

  /**
   * The name for display in the overview section.
   *
   * @readonly
   * @type {string}
   * @memberof MgtProfile
   */
  public get displayName(): string {
    return this.strings.SkillsAndExperienceSectionTitle;
  }

  /**
   * The title for display when rendered as a full card.
   *
   * @readonly
   * @type {string}
   * @memberof MgtOrganization
   */
  public get cardTitle(): string {
    return this.strings.AboutCompactSectionTitle;
  }

  /**
   * Returns true if the profile contains data
   * that can be rendered
   *
   * @readonly
   * @type {boolean}
   * @memberof MgtProfile
   */
  public get hasData(): boolean {
    if (!this.profile) {
      return false;
    }

    var { languages, skills, positions, educationalActivities } = this.profile;

    return (
      [
        this._birthdayAnniversary,
        this._personalInterests?.length,
        this._professionalInterests?.length,
        languages?.length,
        skills?.length,
        positions?.length,
        educationalActivities?.length
      ].filter(v => !!v).length > 0
    );
  }

  /**
   * The user's profile metadata
   *
   * @protected
   * @type {IProfile}
   * @memberof MgtProfile
   */
  protected get profile(): Profile {
    return this._profile;
  }
  @property({ attribute: false })
  protected set profile(value: Profile) {
    if (value === this._profile) {
      return;
    }

    this._profile = value;
    this._birthdayAnniversary = value?.anniversaries ? value.anniversaries.find(this.isBirthdayAnniversary) : null;
    this._personalInterests = value?.interests ? value.interests.filter(this.isPersonalInterest) : null;
    this._professionalInterests = value?.interests ? value.interests.filter(this.isProfessionalInterest) : null;
  }

  private _profile: Profile;
  private _personalInterests: PersonInterest[];
  private _professionalInterests: PersonInterest[];
  private _birthdayAnniversary: PersonAnnualEvent;

  constructor(profile: Profile) {
    super();

    this.profile = profile;
  }

  /**
   * Render the icon for display in the navigation ribbon.
   *
   * @returns {TemplateResult}
   * @memberof MgtProfile
   */
  public renderIcon(): TemplateResult {
    return getSvg(SvgIcon.Profile);
  }

  /**
   * Reset any state in the section
   *
   * @protected
   * @memberof MgtProfile
   */
  public clearState(): void {
    super.clearState();
    this.profile = null;
  }

  /**
   * Render the compact view
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtProfile
   */
  protected renderCompactView(): TemplateResult {
    return html`
       <div class="root compact" dir=${this.direction}>
         ${this.renderSubSections().slice(0, 2)}
       </div>
     `;
  }

  /**
   * Render the full view
   *
   * @protected
   * @returns
   * @memberof MgtProfile
   */
  protected renderFullView() {
    this.initPostRenderOperations();

    return html`
       <div class="root" dir=${this.direction}>
         ${this.renderSubSections()}
       </div>
     `;
  }

  /**
   * Renders all subSections of the profile
   * Defines order of how they render
   *
   * @protected
   * @return {*}
   * @memberof MgtProfile
   */
  protected renderSubSections() {
    var subSections = [
      this.renderSkills(),
      this.renderBirthday(),
      this.renderLanguages(),
      this.renderWorkExperience(),
      this.renderEducation(),
      this.renderProfessionalInterests(),
      this.renderPersonalInterests()
    ];

    return subSections.filter(s => !!s);
  }

  /**
   * Render the user's known languages
   *
   * @protected
   * @returns
   * @memberof MgtProfile
   */
  protected renderLanguages(): TemplateResult {
    var { languages } = this._profile;
    if (!languages?.length) {
      return null;
    }

    var languageItems: TemplateResult[] = [];
    for (var language of languages) {
      let proficiency = null;
      if (language.proficiency?.length) {
        proficiency = html`
           <span class="language__proficiency" tabindex="0">
             &nbsp;(${language.proficiency})
           </span>
         `;
      }

      languageItems.push(html`
         <div class="token-list__item language">
           <span class="language__title" tabindex="0">${language.displayName}</span>
           ${proficiency}
         </div>
       `);
    }

    var languageTitle = languageItems.length ? this.strings.LanguagesSubSectionTitle : '';

    return html`
       <section>
         <div class="section__title" tabindex="0">${languageTitle}</div>
         <div class="section__content">
           <div class="token-list">
             ${languageItems}
           </div>
         </div>
       </section>
     `;
  }

  /**
   * Render the user's skills
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtProfile
   */
  protected renderSkills(): TemplateResult {
    var { skills } = this._profile;

    if (!skills?.length) {
      return null;
    }

    var skillItems: TemplateResult[] = [];
    for (var skill of skills) {
      skillItems.push(html`
         <div class="token-list__item skill" tabindex="0">
           ${skill.displayName}
         </div>
       `);
    }

    var skillsTitle = skillItems.length ? this.strings.SkillsSubSectionTitle : '';

    return html`
       <section>
         <div class="section__title" tabindex="0">${skillsTitle}</div>
         <div class="section__content">
           <div class="token-list">
             ${skillItems}
           </div>
         </div>
       </section>
     `;
  }

  /**
   * Render the user's work experience timeline
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtProfile
   */
  protected renderWorkExperience(): TemplateResult {
    var { positions } = this._profile;

    if (!positions?.length) {
      return null;
    }

    var positionItems: TemplateResult[] = [];
    for (var position of this._profile.positions) {
      if (position.detail.description || position.detail.jobTitle !== '') {
        positionItems.push(html`
           <div class="data-list__item work-position">
             <div class="data-list__item__header">
               <div class="data-list__item__title" tabindex="0">${position.detail?.jobTitle}</div>
               <div class="data-list__item__date-range" tabindex="0">
                 ${this.getDisplayDateRange(position.detail)}
               </div>
             </div>
             <div class="data-list__item__content">
               <div class="work-position__company" tabindex="0">
                 ${position?.detail?.company?.displayName}
               </div>
               <div class="work-position__location" tabindex="0">
                 ${this.displayLocation(position?.detail?.company?.address)}
               </div>
             </div>
           </div>
         `);
      }
    }
    var workExperienceTitle = positionItems.length ? this.strings.WorkExperienceSubSectionTitle : '';

    return html`
       <section>
         <div class="section__title" tabindex="0">${workExperienceTitle}</div>
         <div class="section__content">
           <div class="data-list">
             ${positionItems}
           </div>
         </div>
       </section>
     `;
  }

  /**
   * Render the user's education timeline
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtProfile
   */
  protected renderEducation(): TemplateResult {
    var { educationalActivities } = this._profile;

    if (!educationalActivities?.length) {
      return null;
    }

    var positionItems: TemplateResult[] = [];
    for (var educationalActivity of educationalActivities) {
      positionItems.push(html`
         <div class="data-list__item educational-activity">
           <div class="data-list__item__header">
             <div class="data-list__item__title" tabindex="0">${educationalActivity.institution.displayName}</div>
             <div class="data-list__item__date-range" tabindex="0">
               ${this.getDisplayDateRange(educationalActivity)}
             </div>
           </div>
           ${
             educationalActivity.program.displayName
               ? html`<div class="data-list__item__content">
                  <div class="educational-activity__degree" tabindex="0">
                  ${educationalActivity.program.displayName}
                </div>`
               : nothing
           }
         </div>
       `);
    }

    var educationTitle = positionItems.length ? this.strings.EducationSubSectionTitle : '';

    return html`
       <section>
         <div class="section__title" tabindex="0">${educationTitle}</div>
         <div class="section__content">
           <div class="data-list">
             ${positionItems}
           </div>
         </div>
       </section>
     `;
  }

  /**
   * Render the user's professional interests
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtProfile
   */
  protected renderProfessionalInterests(): TemplateResult {
    if (!this._professionalInterests?.length) {
      return null;
    }

    var interestItems: TemplateResult[] = [];
    for (var interest of this._professionalInterests) {
      interestItems.push(html`
         <div class="token-list__item interest interest--professional" tabindex="0">
           ${interest.displayName}
         </div>
       `);
    }

    var professionalInterests = interestItems.length ? this.strings.professionalInterestsSubSectionTitle : '';

    return html`
       <section>
         <div class="section__title" tabindex="0">${professionalInterests}</div>
         <div class="section__content">
           <div class="token-list">
             ${interestItems}
           </div>
         </div>
       </section>
     `;
  }

  /**
   * Render the user's personal interests
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtProfile
   */
  protected renderPersonalInterests(): TemplateResult {
    if (!this._personalInterests?.length) {
      return null;
    }

    var interestItems: TemplateResult[] = [];
    for (var interest of this._personalInterests) {
      interestItems.push(html`
         <div class="token-list__item interest interest--personal" tabindex="0">
           ${interest.displayName}
         </div>
       `);
    }

    var personalInterests = interestItems.length ? this.strings.personalInterestsSubSectionTitle : '';

    return html`
       <section>
         <div class="section__title" tabindex="0">${personalInterests}</div>
         <div class="section__content">
           <div class="token-list">
             ${interestItems}
           </div>
         </div>
       </section>
     `;
  }

  /**
   * Render the user's birthday
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtProfile
   */
  protected renderBirthday(): TemplateResult {
    if (!this._birthdayAnniversary?.date) {
      return null;
    }

    return html`
       <section>
         <div class="section__title" tabindex="0">Birthday</div>
         <div class="section__content">
           <div class="birthday">
             <div class="birthday__icon">
               ${getSvg(SvgIcon.Birthday)}
             </div>
             <div class="birthday__date" tabindex="0">
               ${this.getDisplayDate(new Date(this._birthdayAnniversary.date))}
             </div>
           </div>
         </div>
       </section>
     `;
  }

  private readonly isPersonalInterest = (interest: PersonInterest): boolean => {
    return interest.categories?.includes('personal');
  };

  private readonly isProfessionalInterest = (interest: PersonInterest): boolean => {
    return interest.categories?.includes('professional');
  };

  private readonly isBirthdayAnniversary = (anniversary: PersonAnnualEvent): boolean => {
    return anniversary.type === 'birthday';
  };

  private getDisplayDate(date: Date): string {
    return date.toLocaleString('default', {
      day: 'numeric',
      month: 'long'
    });
  }

  private getDisplayDateRange(event: EducationalActivity): string | symbol {
    // if startMonthYear is not defined, we do not show the date range (otherwise it will always start with 1970)
    if (!event.startMonthYear) {
      return nothing;
    }

    var start = new Date(event.startMonthYear).getFullYear();
    // if the start year is 0 or 1 - it's probably an error or a strange "undefined"-value
    if (start === 0 || start === 1) {
      return nothing;
    }

    var end = event.endMonthYear ? new Date(event.endMonthYear).getFullYear() : this.strings.currentYearSubtitle;
    return `${start} — ${end}`;
  }

  private displayLocation(address: PhysicalAddress | undefined): string | symbol {
    if (address?.city) {
      if (address.state) {
        return `${address.city}, ${address.state}`;
      }
      return address.city;
    }
    return nothing;
  }

  private initPostRenderOperations(): void {
    setTimeout(() => {
      try {
        var sections = this.shadowRoot.querySelectorAll('section');
        sections.forEach(section => {
          // Perform post render operations per section
          this.handleTokenOverflow(section);
        });
      } catch {
        // An exception may occur if the component is suddenly removed during post render operations.
      }
    }, 0);
  }

  private handleTokenOverflow(section: HTMLElement): void {
    var tokenLists = section.querySelectorAll('.token-list');
    if (!tokenLists?.length) {
      return;
    }

    for (var tokenList of Array.from(tokenLists)) {
      var items = tokenList.querySelectorAll('.token-list__item');
      if (!items?.length) {
        continue;
      }

      let overflowItems: Element[] = null;
      let itemRect = items[0].getBoundingClientRect();
      var tokenListRect = tokenList.getBoundingClientRect();
      var maxtop = itemRect.height * 2 + tokenListRect.top;

      // Use (items.length - 1) to prevent [+1 more] from appearing.
      for (let i = 0; i < items.length - 1; i++) {
        itemRect = items[i].getBoundingClientRect();
        if (itemRect.top > maxtop) {
          overflowItems = Array.from(items).slice(i, items.length);
          break;
        }
      }

      if (overflowItems) {
        overflowItems.forEach(i => i.classList.add('overflow'));

        var overflowToken = document.createElement('div');
        overflowToken.classList.add('token-list__item');
        overflowToken.classList.add('token-list__item--show-overflow');
        overflowToken.tabIndex = 0;
        overflowToken.innerText = `+ ${overflowItems.length} more`;

        // On click or enter(accessibility), remove [+n more] token and reveal the hidden overflow tokens.
        var revealOverflow = () => {
          overflowToken.remove();
          overflowItems.forEach(i => i.classList.remove('overflow'));
        };
        overflowToken.addEventListener('click', () => {
          revealOverflow();
        });
        overflowToken.addEventListener('keydown', (e: KeyboardEvent) => {
          if (e.code === 'Enter') {
            revealOverflow();
          }
        });
        tokenList.appendChild(overflowToken);
      }
    }
  }
}
