/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */


"use strict";

/**
 * Sections of a presentation (PowerPoint's p14:sectionLst). The first slide of a section carries the header
 * (slide.sectionId, slide.sectionName); the other slides belong to the section by their position, so the
 * ordinary slide commands (new slide, delete, move by drag and drop, undo) keep sections consistent:
 *  - deleting or moving the first slide of a section hands the header to the next slide of the section,
 *  - a slide moved to the front joins the first section,
 *  - a section whose slides are all gone disappears.
 * If any slide has a header the first slide has one too (the "Default Section" Microsoft Office also adds).
 * The headers are saved in the editor binary (record 7 of the presentation, common/Shapes/SerializeWriter.js,
 * read in common/Drawings/Format/Format.js); the core converter (PPTX::Presentation) writes the slides of a
 * section as their positions in the list of slides and gives them the ids of the file.
 *   asc_GetSections, asc_AddSection, asc_RenameSection, asc_RemoveSection, asc_RemoveSectionAndSlides,
 *   asc_RemoveAllSections, asc_MoveSection, asc_SetSectionCollapsed, asc_SelectSection; event asc_onSectionsChanged
 */
(function (window) {
	const CPresentation = window["AscCommonSlide"].CPresentation;
	const Slide = window["AscCommonSlide"].Slide;
	const api = window["Asc"]["asc_docs_api"];
	const History = AscCommon.History;
	const AscDFH = window["AscDFH"];

	const DEFAULT_NAME = "Default Section";
	const NEW_NAME = "Untitled Section";

	AscDFH.historyitem_SlideSetSectionId = AscDFH.historyitem_type_Slide | 17;
	AscDFH.historyitem_SlideSetSectionName = AscDFH.historyitem_type_Slide | 18;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetSectionId] = AscDFH.CChangesDrawingsString;
	AscDFH.changesFactory[AscDFH.historyitem_SlideSetSectionName] = AscDFH.CChangesDrawingsString;
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetSectionId] = function (oClass, value) {
		oClass.sectionId = value || null;
	};
	AscDFH.drawingsChangesMap[AscDFH.historyitem_SlideSetSectionName] = function (oClass, value) {
		oClass.sectionName = value || null;
	};

	Slide.prototype.sectionId = null;
	Slide.prototype.sectionName = null;

	Slide.prototype.hasSectionHeader = function () {
		return !!this.sectionId;
	};
	/** Sets the header (id and name) of the section this slide starts, or none; recorded in the history */
	Slide.prototype.setSectionHeader = function (id, name) {
		id = id || null;
		name = id ? (name || NEW_NAME) : null;
		if (this.sectionId !== id) {
			History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_SlideSetSectionId, this.sectionId, id));
			this.sectionId = id;
		}
		if (this.sectionName !== name) {
			History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_SlideSetSectionName, this.sectionName, name));
			this.sectionName = name;
		}
	};

	function newId() {
		return AscCommon.CreateGUID();
	}

	// ---- keeping sections consistent when slides come and go ------------------------------------
	const originalRemove = CPresentation.prototype.removeSlideByObject;
	CPresentation.prototype.removeSlideByObject = function (oSlide, bNoCheck, pos) {
		if (oSlide && oSlide.getObjectType && oSlide.getObjectType() === AscDFH.historyitem_type_Slide && oSlide.hasSectionHeader()) {
			const index = typeof pos === "number" ? pos : this.Slides.indexOf(oSlide);
			const next = this.Slides[index + 1];
			// the section goes on with its next slide, unless that slide starts another section
			if (next && !next.hasSectionHeader()) {
				next.setSectionHeader(oSlide.sectionId, oSlide.sectionName);
			}
			oSlide.setSectionHeader(null);
		}
		return originalRemove.apply(this, arguments);
	};

	const originalInsert = CPresentation.prototype.insertSlide;
	CPresentation.prototype.insertSlide = function (pos, slide) {
		const res = originalInsert.apply(this, arguments);
		// a slide put in front of the first section joins it
		const next = this.Slides[pos + 1];
		if (pos === 0 && next && next.hasSectionHeader() && slide.getObjectType() === AscDFH.historyitem_type_Slide && !slide.hasSectionHeader()) {
			slide.setSectionHeader(next.sectionId, next.sectionName);
			next.setSectionHeader(null);
		}
		return res;
	};

	// ---- the model --------------------------------------------------------------------------------
	/**
	 * @returns {Array<{id: string, name: string, first: number, count: number, collapsed: boolean}>} empty without sections
	 */
	CPresentation.prototype.GetSections = function () {
		const slides = this.Slides, res = [];
		if (!slides.some(function (s) {
			return s.hasSectionHeader();
		})) {
			return res;
		}
		const collapsed = this.elixirCollapsedSections || {};
		let cur = null;
		for (let i = 0; i < slides.length; ++i) {
			if (i === 0 || slides[i].hasSectionHeader()) {
				cur = {id: slides[i].sectionId || "", name: slides[i].sectionName || DEFAULT_NAME, first: i, count: 0, collapsed: false};
				cur.collapsed = !!collapsed[cur.id];
				res.push(cur);
			}
			cur.count++;
		}
		return res;
	};

	/** The slides of the section, by index */
	function sectionSlides(section) {
		const list = [];
		for (let i = section.first; i < section.first + section.count; ++i) {
			list.push(i);
		}
		return list;
	}

	function findSection(presentation, id) {
		const sections = presentation.GetSections();
		for (let i = 0; i < sections.length; ++i) {
			if (sections[i].id === id) {
				return {sections: sections, index: i, section: sections[i]};
			}
		}
		return null;
	}

	function run(presentation, slides, fn) {
		if (!presentation.CanEdit() || presentation.IsMasterMode() || presentation.Document_Is_SelectionLocked(AscCommon.changestype_SlideBg, slides) !== false) {
			return false;
		}
		presentation.StartAction(AscDFH.historydescription_Presentation_ShiftSlides);
		const res = fn();
		presentation.FinalizeAction();
		presentation.DrawingDocument.UpdateThumbnailsAttack();
		presentation.Document_UpdateInterfaceState();
		notify(presentation);
		return res === undefined ? true : res;
	}

	/**
	 * Adds a section starting at a slide (the first selected slide by default). Without sections the slides before it
	 * become the "Default Section", as in PowerPoint.
	 * @returns {string} the id of the new section, or "" when nothing was added
	 */
	CPresentation.prototype.AddSection = function (index) {
		if (index === undefined || index === null) {
			const selected = this.GetSelectedSlides();
			index = selected.length ? selected[0] : this.CurPage;
		}
		const presentation = this, slide = this.Slides[index];
		if (!slide || slide.hasSectionHeader()) {
			return "";
		}
		const touched = [slide];
		if (!this.Slides[0].hasSectionHeader()) {
			touched.push(this.Slides[0]);
		}
		const id = newId();
		return run(this, touched, function () {
			if (index > 0 && !presentation.Slides[0].hasSectionHeader()) {
				presentation.Slides[0].setSectionHeader(newId(), DEFAULT_NAME);
			}
			slide.setSectionHeader(id, NEW_NAME);
			return true;
		}) ? id : "";
	};

	CPresentation.prototype.RenameSection = function (id, name) {
		const found = findSection(this, id);
		name = (name || "").trim();
		if (!found || !name || found.section.name === name) {
			return false;
		}
		const slide = this.Slides[found.section.first];
		return run(this, [slide], function () {
			slide.setSectionHeader(slide.sectionId, name);
		});
	};

	/** Removes the section, its slides join the previous one (the next one for the first section) */
	CPresentation.prototype.RemoveSection = function (id) {
		const found = findSection(this, id);
		if (!found) {
			return false;
		}
		const presentation = this, s = found.section, slide = this.Slides[s.first];
		const next = found.sections[found.index + 1];
		const touched = [slide];
		next && touched.push(this.Slides[next.first]);
		return run(this, touched, function () {
			if (found.index === 0 && next) {
				// the first section goes: the next one starts at the first slide
				slide.setSectionHeader(next.id, next.name);
				presentation.Slides[next.first].setSectionHeader(null);
			} else {
				slide.setSectionHeader(null);
			}
		});
	};

	CPresentation.prototype.RemoveSectionAndSlides = function (id) {
		const found = findSection(this, id);
		if (!found || !this.CanEdit()) {
			return false;
		}
		const slides = sectionSlides(found.section).map(function (i) {
			return this.Slides[i];
		}, this);
		this.deleteSlides(slides);
		notify(this);
		return true;
	};

	CPresentation.prototype.RemoveAllSections = function () {
		const slides = this.Slides.filter(function (s) {
			return s.hasSectionHeader();
		});
		if (!slides.length) {
			return false;
		}
		const presentation = this;
		return run(this, slides, function () {
			slides.forEach(function (s) {
				s.setSectionHeader(null);
			});
			presentation.elixirCollapsedSections = {};
		});
	};

	/** Moves a section up (direction -1) or down (+1) among the others, with its slides */
	CPresentation.prototype.MoveSection = function (id, direction) {
		const found = findSection(this, id);
		if (!found) {
			return false;
		}
		const other = found.sections[found.index + (direction < 0 ? -1 : 1)];
		if (!other) {
			return false;
		}
		const presentation = this, s = found.section;
		const objects = function (sec) {
			return sectionSlides(sec).map(function (i) {
				return presentation.Slides[i];
			});
		};
		const mine = objects(s), theirs = objects(other);
		const order = direction < 0 ? [{sec: s, slides: mine}, {sec: other, slides: theirs}] : [{sec: other, slides: theirs}, {sec: s, slides: mine}];
		return run(this, mine.concat(theirs), function () {
			// the slides of this section go before / after the other one (positions after the removal of the moved slides)
			const indexes = sectionSlides(s);
			presentation.moveSlides(indexes, direction < 0 ? other.first : s.first + other.count);
			// the headers went with the individual slides: give each section its header back, on its first slide
			order.forEach(function (o) {
				o.slides.forEach(function (sl, k) {
					if (k === 0) {
						sl.setSectionHeader(o.sec.id || newId(), o.sec.name);
					} else {
						sl.setSectionHeader(null);
					}
				});
			});
			presentation.updateSlideIndexes && presentation.updateSlideIndexes();
			presentation.Recalculate();
		});
	};

	CPresentation.prototype.SetSectionCollapsed = function (id, isCollapsed) {
		this.elixirCollapsedSections = this.elixirCollapsedSections || {};
		if (isCollapsed) {
			this.elixirCollapsedSections[id] = true;
		} else {
			delete this.elixirCollapsedSections[id];
		}
		this.DrawingDocument.UpdateThumbnailsAttack();
		notify(this, true);
	};

	CPresentation.prototype.SelectSection = function (id) {
		const found = findSection(this, id);
		if (!found) {
			return false;
		}
		const thumbnails = this.Api.WordControl && this.Api.WordControl.Thumbnails;
		const list = sectionSlides(found.section);
		this.DrawingDocument.m_oWordControl.GoToPage(list[0]);
		thumbnails && thumbnails.SelectSlides && thumbnails.SelectSlides(list);
		return true;
	};

	// ---- load / save --------------------------------------------------------------------------------
	/** From the file: [{id, name, slides: [position, ...]}]; no history */
	CPresentation.prototype.applyLoadedSections = function (list) {
		const slides = this.Slides;
		(list || []).forEach(function (section) {
			const positions = (section.slides || []).filter(function (p) {
				return p >= 0 && p < slides.length;
			});
			if (!positions.length) {
				return;
			}
			const slide = slides[Math.min.apply(null, positions)];
			slide.sectionId = section.id || newId();
			slide.sectionName = section.name || NEW_NAME;
		});
		if (slides.length && !slides[0].hasSectionHeader() && slides.some(function (s) {
			return s.hasSectionHeader();
		})) {
			slides[0].sectionId = newId();
			slides[0].sectionName = DEFAULT_NAME;
		}
	};

	/** For the file: [{id, name, slides: [position, ...]}], none when there are no sections */
	CPresentation.prototype.getSectionsToSave = function () {
		return this.GetSections().map(function (s) {
			return {id: s.id || newId(), name: s.name, slides: sectionSlides(s)};
		});
	};

	// ---- notification ---------------------------------------------------------------------------------
	function signature(presentation) {
		return presentation.GetSections().map(function (s) {
			return [s.id, s.name, s.first, s.count, s.collapsed ? 1 : 0].join("|");
		}).join(";");
	}

	function notify(presentation, force) {
		const sig = signature(presentation);
		if (force || sig !== presentation.elixirSectionsSignature) {
			presentation.elixirSectionsSignature = sig;
			presentation.Api.sendEvent("asc_onSectionsChanged", presentation.GetSections());
		}
	}

	["Document_UpdateInterfaceState", "Document_UpdateUndoRedoState"].forEach(function (name) {
		const original = CPresentation.prototype[name];
		CPresentation.prototype[name] = function () {
			const res = original.apply(this, arguments);
			if (signature(this) !== this.elixirSectionsSignature) {
				this.DrawingDocument.UpdateThumbnailsAttack && this.DrawingDocument.UpdateThumbnailsAttack();
				notify(this);
			}
			return res;
		};
	});

	// going to a slide of a collapsed section opens the section
	const originalSetCurPage = CPresentation.prototype.Set_CurPage;
	CPresentation.prototype.Set_CurPage = function (index) {
		const res = originalSetCurPage.apply(this, arguments);
		const collapsed = this.elixirCollapsedSections;
		if (collapsed && Object.keys(collapsed).length && this.Slides[this.CurPage]) {
			const sections = this.GetSections();
			for (let i = 0; i < sections.length; ++i) {
				const s = sections[i];
				if (s.collapsed && this.CurPage >= s.first && this.CurPage < s.first + s.count) {
					this.SetSectionCollapsed(s.id, false);
				}
			}
		}
		return res;
	};

	// ---- API ---------------------------------------------------------------------------------------------
	function presentationOf(editor) {
		return editor.WordControl && editor.WordControl.m_oLogicDocument;
	}
	api.prototype["asc_GetSections"] = api.prototype.asc_GetSections = function () {
		const p = presentationOf(this);
		return p ? p.GetSections() : [];
	};
	api.prototype["asc_AddSection"] = api.prototype.asc_AddSection = function (index) {
		const p = presentationOf(this);
		return p ? p.AddSection(index) : "";
	};
	api.prototype["asc_RenameSection"] = api.prototype.asc_RenameSection = function (id, name) {
		const p = presentationOf(this);
		return p ? p.RenameSection(id, name) : false;
	};
	api.prototype["asc_RemoveSection"] = api.prototype.asc_RemoveSection = function (id) {
		const p = presentationOf(this);
		return p ? p.RemoveSection(id) : false;
	};
	api.prototype["asc_RemoveSectionAndSlides"] = api.prototype.asc_RemoveSectionAndSlides = function (id) {
		const p = presentationOf(this);
		return p ? p.RemoveSectionAndSlides(id) : false;
	};
	api.prototype["asc_RemoveAllSections"] = api.prototype.asc_RemoveAllSections = function () {
		const p = presentationOf(this);
		return p ? p.RemoveAllSections() : false;
	};
	api.prototype["asc_MoveSection"] = api.prototype.asc_MoveSection = function (id, direction) {
		const p = presentationOf(this);
		return p ? p.MoveSection(id, direction) : false;
	};
	api.prototype["asc_SetSectionCollapsed"] = api.prototype.asc_SetSectionCollapsed = function (id, isCollapsed) {
		const p = presentationOf(this);
		p && p.SetSectionCollapsed(id, isCollapsed);
	};
	api.prototype["asc_SelectSection"] = api.prototype.asc_SelectSection = function (id) {
		const p = presentationOf(this);
		return p ? p.SelectSection(id) : false;
	};
})(window);
