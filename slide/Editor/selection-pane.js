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
 * Selection pane of PowerPoint (Alt+F10): the objects of the current slide, the one in front first, with
 * select, show / hide (p:cNvPr hidden), rename and z-order. A hidden object stays in the slide, is not
 * drawn (editing, thumbnails, slide show, print) and cannot be hit with the mouse.
 *   asc_GetSelectionPaneObjects, asc_SelectPaneObject, asc_SetPaneObjectsHidden, asc_RenamePaneObject
 */
(function (window) {
	const CPresentation = window["AscCommonSlide"].CPresentation;
	const api = window["Asc"]["asc_docs_api"];
	const AscFormat = window["AscFormat"];

	function isHidden(drawing) {
		const props = drawing && drawing.getCNvProps && drawing.getCNvProps();
		return !!(props && props.isHidden);
	}

	// ---- a hidden object is neither drawn nor hit
	["CShape", "CImageShape", "CGroupShape", "CGraphicFrame", "CChartSpace", "CConnectorShape", "SmartArt"].forEach(function (name) {
		const cls = AscFormat[name];
		if (!cls || !cls.prototype) {
			return;
		}
		const proto = cls.prototype;
		if (typeof proto.draw === "function") {
			const draw = proto.draw;
			proto.draw = function () {
				if (isHidden(this)) {
					return;
				}
				return draw.apply(this, arguments);
			};
		}
		["hit", "hitInInnerArea", "hitInPath", "hitInBoundingRect", "hitInTextRect"].forEach(function (method) {
			if (typeof proto[method] !== "function") {
				return;
			}
			const original = proto[method];
			proto[method] = function () {
				return isHidden(this) ? false : original.apply(this, arguments);
			};
		});
	});

	function slideObjects(presentation) {
		const slide = presentation.GetCurrentSlide();
		return slide && !presentation.IsFocusOnNotes() ? slide.cSld.spTree : [];
	}

	function kindOf(drawing) {
		switch (drawing.getObjectType()) {
			case AscDFH.historyitem_type_ImageShape:
				return "picture";
			case AscDFH.historyitem_type_GroupShape:
				return "group";
			case AscDFH.historyitem_type_ChartSpace:
				return "chart";
			case AscDFH.historyitem_type_GraphicFrame:
				return drawing.graphicObject && drawing.graphicObject.Content ? "table" : "object";
			case AscDFH.historyitem_type_SmartArt:
				return "smartart";
			default:
				return drawing.isPlaceholder && drawing.isPlaceholder() ? "placeholder" : "shape";
		}
	}

	CPresentation.prototype.GetSelectionPaneObjects = function () {
		const controller = this.GetCurrentController();
		const selected = controller ? controller.selectedObjects : [];
		const list = slideObjects(this).map(function (drawing, index) {
			return {
				"id": drawing.Get_Id(),
				"name": drawing.getObjectName(),
				"kind": kindOf(drawing),
				"hidden": isHidden(drawing),
				"selected": selected.indexOf(drawing) !== -1,
				"z": index
			};
		});
		// the object in front (the last of the slide) comes first
		return list.reverse();
	};

	function paneObject(presentation, id) {
		const drawing = AscCommon.g_oTableId.Get_ById(id);
		return drawing && slideObjects(presentation).indexOf(drawing) !== -1 ? drawing : null;
	}

	CPresentation.prototype.SelectPaneObject = function (id, isAdd) {
		const drawing = paneObject(this, id), controller = this.GetCurrentController();
		if (!drawing || !controller || isHidden(drawing)) {
			return false;
		}
		this.StopAnimationPreview && this.StopAnimationPreview();
		if (!isAdd) {
			controller.resetSelection();
		}
		drawing.select(controller, this.CurPage);
		this.Document_UpdateSelectionState();
		this.Document_UpdateInterfaceState();
		return true;
	};

	/**
	 * Show or hide objects of the current slide (one undo step); ids: one id or a list, or null for all of them
	 */
	CPresentation.prototype.SetPaneObjectsHidden = function (ids, isHidden_) {
		const presentation = this, controller = this.GetCurrentController();
		if (!controller || !this.CanEdit()) {
			return;
		}
		let drawings = null === ids ? slideObjects(this).slice() : [].concat(ids).map(function (id) {
			return paneObject(presentation, id);
		}).filter(Boolean);
		drawings = drawings.filter(function (drawing) {
			return isHidden(drawing) !== isHidden_;
		});
		if (!drawings.length) {
			return;
		}
		controller.checkSelectedObjectsAndCallback(function () {
			if (isHidden_) {
				const selected = controller.selectedObjects.slice();
				if (drawings.some(function (drawing) {
					return selected.indexOf(drawing) !== -1;
				})) {
					controller.resetSelection();
				}
			}
			drawings.forEach(function (drawing) {
				drawing.getCNvProps().setIsHidden(isHidden_ ? true : null);
			});
		}, [], false, AscDFH.historydescription_CommonControllerCheckSelected, drawings);
		this.Document_UpdateSelectionState();
		this.Document_UpdateInterfaceState();
		this.DrawingDocument.OnRecalculateSlide(this.CurPage);
		this.DrawingDocument.UpdateThumbnailsAttack();
	};

	CPresentation.prototype.RenamePaneObject = function (id, name) {
		const drawing = paneObject(this, id), controller = this.GetCurrentController();
		if (!drawing || !name || !controller || !this.CanEdit()) {
			return;
		}
		controller.checkSelectedObjectsAndCallback(function () {
			drawing.getCNvProps().setName(name);
		}, [], false, AscDFH.historydescription_CommonControllerCheckSelected, [drawing]);
	};

	/** Bring Forward / Send Backward (one step) for the objects selected in the pane */
	CPresentation.prototype.MovePaneObject = function (forward) {
		forward ? this.bringForward() : this.bringBackward();
	};

	api.prototype["asc_GetSelectionPaneObjects"] = api.prototype.asc_GetSelectionPaneObjects = function () {
		const presentation = this.WordControl && this.WordControl.m_oLogicDocument;
		return presentation ? presentation.GetSelectionPaneObjects() : [];
	};
	api.prototype["asc_SelectPaneObject"] = api.prototype.asc_SelectPaneObject = function (id, isAdd) {
		const presentation = this.WordControl && this.WordControl.m_oLogicDocument;
		return presentation ? presentation.SelectPaneObject(id, isAdd) : false;
	};
	api.prototype["asc_SetPaneObjectsHidden"] = api.prototype.asc_SetPaneObjectsHidden = function (ids, isHidden_) {
		const presentation = this.WordControl && this.WordControl.m_oLogicDocument;
		presentation && presentation.SetPaneObjectsHidden(ids, isHidden_);
	};
	api.prototype["asc_RenamePaneObject"] = api.prototype.asc_RenamePaneObject = function (id, name) {
		const presentation = this.WordControl && this.WordControl.m_oLogicDocument;
		presentation && presentation.RenamePaneObject(id, name);
	};
	api.prototype["asc_MovePaneObject"] = api.prototype.asc_MovePaneObject = function (forward) {
		const presentation = this.WordControl && this.WordControl.m_oLogicDocument;
		presentation && presentation.MovePaneObject(forward);
	};
})(window);
