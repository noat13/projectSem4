package vn.com.unit.coteccons.controller;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.UUID;

import javax.servlet.ServletContext;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.validation.Valid;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.BooleanUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.joda.time.DateTime;
import org.joda.time.DateTimeFieldType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.propertyeditors.CustomDateEditor;
import org.springframework.beans.propertyeditors.CustomNumberEditor;
import org.springframework.context.ApplicationContext;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.stereotype.Controller;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.InitBinder;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.support.ByteArrayMultipartFileEditor;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import vn.com.unit.binding.DoubleEditor;
import vn.com.unit.coteccons.bean.BillBean;
import vn.com.unit.coteccons.bean.EquipmentCategoryTreeBean;
import vn.com.unit.coteccons.bean.EquipmentPriceBean;
import vn.com.unit.coteccons.bean.Message;
import vn.com.unit.coteccons.bean.UserProfile;
import vn.com.unit.coteccons.config.SystemConfig;
import vn.com.unit.coteccons.entity.Account;
import vn.com.unit.coteccons.entity.Bill;
import vn.com.unit.coteccons.entity.BillDetail;
import vn.com.unit.coteccons.entity.Equipment;
import vn.com.unit.coteccons.entity.EquipmentCategory;
import vn.com.unit.coteccons.entity.HistoryApprove;
import vn.com.unit.coteccons.entity.ProcessStep;
import vn.com.unit.coteccons.entity.Project;
import vn.com.unit.coteccons.entity.Stock;
import vn.com.unit.coteccons.entity.StockTracking;
import vn.com.unit.coteccons.scheduler.BillTask;
import vn.com.unit.coteccons.service.AccountProjectService;
import vn.com.unit.coteccons.service.AccountService;
import vn.com.unit.coteccons.service.BillDetailService;
import vn.com.unit.coteccons.service.BillService;
import vn.com.unit.coteccons.service.EquipmentCategoryService;
import vn.com.unit.coteccons.service.HistoryApproveService;
import vn.com.unit.coteccons.service.PriceProjectSettingService;
import vn.com.unit.coteccons.service.ProcessStepPendingService;
import vn.com.unit.coteccons.service.ProcessStepService;
import vn.com.unit.coteccons.service.ProjectService;
import vn.com.unit.coteccons.service.StockService;
import vn.com.unit.coteccons.service.StockTrackingService;
import vn.com.unit.coteccons.utils.ExcelHelper;
import vn.com.unit.coteccons.utils.Utils;
import vn.com.unit.coteccons.utils.ajax.ReturnObject;
import vn.com.unit.coteccons.utils.excel.PoiToHtmlConverter;

/**
 * 
 * @author CongDT
 * @since Mar 4, 2015 11:19:07 AM
 * 
 */
@Controller
@RequestMapping(value = "/Bill")
public class BillController {

	@Autowired
	MessageSource msgSrc;

	@Autowired
	UserProfile userProfile;

	@Autowired
	ServletContext servletContext;
	@Autowired
	SystemConfig systemConfig;

	@Autowired
	StockService stockService;

	@Autowired
	ProjectService projectService;

	@Autowired
	BillService billService;

	@Autowired
	BillDetailService billDetailService;

	@Autowired
	ProcessStepService processStepService;

	@Autowired
	AccountService accountService;

	@Autowired
	AccountProjectService accountProjectService;

	@Autowired
	private ApplicationContext appContext;

	@Autowired
	ProcessStepPendingService processStepPendingService;

	@Autowired
	EquipmentCategoryService equipmentCategoryService;

	@Autowired
	HistoryApproveService historyApproveService;

	@Autowired
	StockTrackingService stockTrackingService;

	@Autowired
	PriceProjectSettingService priceProjectSettingService;

	private static final Logger logger = LoggerFactory.getLogger(BillController.class);

	@InitBinder
	public void initBinderData(WebDataBinder binder, Locale locale, HttpServletRequest request) {

		// Binder Date
		binder.setAutoGrowCollectionLimit(10000);
		// The date format to parse or output your dates
		SimpleDateFormat dateFormat = new SimpleDateFormat((String) request.getSession().getAttribute("formatDate"));
		// Create a new CustomDateEditor
		CustomDateEditor customDateEditor = new CustomDateEditor(dateFormat, true);
		// Register it as custom editor for the Date type
		binder.registerCustomEditor(Date.class, customDateEditor);

		dateFormat = new SimpleDateFormat("MM/yyyy");
		customDateEditor = new CustomDateEditor(dateFormat, true);
		binder.registerCustomEditor(Date.class, "billPeriod", customDateEditor);
		binder.registerCustomEditor(Date.class, "billFrom", customDateEditor);
		binder.registerCustomEditor(Date.class, "billTo", customDateEditor);

		// Binder BigDecimal value
		NumberFormat numberFormat = NumberFormat.getInstance(locale);
		CustomNumberEditor customNumberEditor = new CustomNumberEditor(BigDecimal.class, numberFormat, true);
		binder.registerCustomEditor(BigDecimal.class, customNumberEditor);

		// Binder Double value
		binder.registerCustomEditor(byte[].class, new ByteArrayMultipartFileEditor());

		binder.registerCustomEditor(Double.class, new DoubleEditor(locale, "#,###.##"));
	}

	@RequestMapping(value = "/json_parent_combo", method = RequestMethod.GET)
	public @ResponseBody List<EquipmentCategoryTreeBean> json_parent_combo(@RequestParam Long currentNode, Locale locale) {
		List<EquipmentCategoryTreeBean> equipmentCategoryTree = equipmentCategoryService.createTreeByEquipmentCategoryId(currentNode);
		return equipmentCategoryTree;
	}

	@RequestMapping(value = "/listStock", method = { RequestMethod.GET, RequestMethod.POST })
	public String listSaveSessionBillStock(@ModelAttribute(value = "bean") BillBean bean, Model model, Locale locale, HttpServletRequest request) {
		bean.clearMessages();
		try {
			Account user = userProfile.getAccount();
			// User thuộc kho nào thì chỉ hiển thị bill của kho đó
			// R019 Bảng chi phí quản lý (Admin)
			// R020 Bảng chi phí kho
			// R021 Bảng chi phí công trường
			if (request.isUserInRole("R019")) {
				// cho xem hết bill
			} else if (request.isUserInRole("R020")) {
				// Kiểm tra user thuộc kho nào
				if (user.getStock() == null || user.getStock().getStockId() == null) {
					throw new Exception(getMsg("Bill.msg.NotBelongToStock"));
				}
				bean.setStockId(user.getStock().getStockId());
				model.addAttribute("IsLockStock", true);
			} else {
				throw new Exception(getMsg("Bill.msg.AccessNotPermitted"));
			}

			doListBillGeneral(bean, model, locale, request);

		} catch (Exception e) {
			logger.debug("##listStock##", e);
			bean.addMessage(Message.ERROR, String.valueOf(e.getMessage()));
		}

		return "Bill.list";
	}

	@RequestMapping(value = "/listProject", method = { RequestMethod.GET, RequestMethod.POST })
	public String listSaveSessionBillProject(@ModelAttribute(value = "bean") BillBean bean, Model model, Locale locale, HttpServletRequest request) {
		bean.clearMessages();
		try {
			Account user = userProfile.getAccount();
			// User thuộc kho nào thì chỉ hiển thị bill của kho đó
			// R019 Bảng chi phí quản lý (Admin)
			// R020 Bảng chi phí kho
			// R021 Bảng chi phí công trường
			if (request.isUserInRole("R019")) {
				// Admin
			} else if (request.isUserInRole("R021")) {
				// Kiểm tra xem user thuộc công trường nào
				List<Project> projects = accountProjectService.getProjectByAccount(user.getId(), false);
				/*
				 * if (user.getProject() == null || user.getProject().getProjectId() == null) { throw new
				 * Exception(getMsg("Bill.msg.NotBelongToProject")); }
				 */

				if (projects == null || projects.size() == 0) {
					throw new Exception(getMsg("Bill.msg.NotBelongToProject"));
				}

				bean.setStockId(user.getStock().getStockId());
				// bean.setProjectId(user.getProject().getProjectId());

				List<Long> longs = new ArrayList<Long>();
				for (Project project : projects) {
					longs.add(project.getProjectId());
				}
				bean.setProjectIds(longs);

				model.addAttribute("IsLockStock", true);
				model.addAttribute("IsLockProject", true);
			} else {
				throw new Exception(getMsg("Bill.msg.AccessNotPermitted"));
			}

			doListBillGeneral(bean, model, locale, request);
		} catch (Exception e) {
			logger.debug("##listProject##", e);
			bean.addMessage(Message.ERROR, String.valueOf(e.getMessage()));
		}

		return "Bill.list";
	}

	private void doListBillGeneral(@ModelAttribute(value = "bean") BillBean bean, Model model, Locale locale, HttpServletRequest request)
			throws Exception {

		bean.setPendingAt(userProfile.getAccount().getUsername());

		// Limit dữ liệu mỗi trang
		int pagesize = Integer.parseInt(systemConfig.getConfig(SystemConfig.PAGING_SIZE));
		bean.setLimit(pagesize);

		// Sắp xếp dữ liệu ASCENDING theo priceTableId
		if (StringUtils.isBlank(bean.getDir())) {
			bean.setSort("desc");
			bean.setDir("billId");
		}

		List<Stock> stocks = stockService.findAll();
		model.addAttribute("stocks", stocks);
		
		List<Project> projects = null;
		
		if (bean.getStockId() != null) {
			projects = projectService.findByStockId(bean.getStockId());			
		} else {
			projects = projectService.findAll();
		}
		model.addAttribute("projects", projects);

		billService.find(bean);

		model.addAttribute("transferStatus", systemConfig.getStatusTransferLinkMap());

//		if (bean.getStockId() != null) {
//			List<Project> projects = projectService.findByStockId(bean.getStockId());
//			model.addAttribute("projects", projects);
//		}

	}

	@RequestMapping(value = "/BillAccumulated", method = { RequestMethod.GET })
	public String doListBillAccumulated(@ModelAttribute(value = "bean") BillBean bean, Model model, Locale locale) {

		try {

			Long projectId = bean.getProjectId();
			if (projectId == null) {
				throw new Exception(msgSrc.getMessage("msg.project.null", null, locale));
			}

			Project project = projectService.findById(projectId);
			if (project == null) {
				throw new Exception(msgSrc.getMessage("msg.project.not.found", null, locale));
			}

			String datePattern = systemConfig.getConfigMap().get(SystemConfig.DATE_PATTERN);

			String projectName = project.getName();
			String stockName = project.getStock().getName();
			String projectStartDate = new SimpleDateFormat(datePattern).format(project.getStartDate());

			model.addAttribute("projectName", projectName);
			model.addAttribute("stockName", stockName);
			model.addAttribute("projectStartDate", projectStartDate);

			String billPeriodTo = null;
			if (bean.getBillPeriod() != null) {
				billPeriodTo = new SimpleDateFormat("yyyyMM").format(bean.getBillPeriod());
			}

			List<EquipmentCategory> roots = equipmentCategoryService.getEquipmentCategoryRoots();
			if (CollectionUtils.isEmpty(roots)) {
				throw new Exception(msgSrc.getMessage("equipment.root.not.found", null, locale));
			}

			// Cột động
			Map<String, String> colNameMap = new HashMap<String, String>();
			Set<String> dynamicCol = new LinkedHashSet<String>();
			for (EquipmentCategory root : roots) {
				dynamicCol.add(root.getEquipmentCategoryCode());
				colNameMap.put(root.getEquipmentCategoryCode(), root.getName());
			}
			model.addAttribute("colNameMap", colNameMap);
			model.addAttribute("dynamicCol", dynamicCol);

			List<Map<String, Object>> billAccumulatedList = billService.getBillAccumulatedList(projectId, billPeriodTo, dynamicCol);
			model.addAttribute("billAccumulatedList", billAccumulatedList);

		} catch (Exception e) {
			bean.addMessage(Message.ERROR, String.valueOf(e.getMessage()));
		}

		return "Bill.list.Accumulated";
	}

	@RequestMapping(value = "/ReportBillAccumulated", method = { RequestMethod.GET })
	public String doReportBillAccumulatedGet(@ModelAttribute(value = "bean") BillBean bean, Model model, Locale locale) {

		try {

			List<Stock> stocks = stockService.findAll();
			model.addAttribute("stocks", stocks);

			List<Project> projects = projectService.findAll();
			model.addAttribute("projects", projects);

			bean.setIsIncludeInternal(true);

		} catch (Exception e) {
			bean.addMessage(Message.ERROR, String.valueOf(e.getMessage()));
		}

		return "Bill.Report.BillAccumulated";

	}

	@RequestMapping(value = "/quickPrint", method = { RequestMethod.GET })
	@ResponseBody
	public Object doQuickPrint(@RequestParam(value = "billId", required = true) long billId, Model model, Locale locale, HttpServletRequest req,
			HttpServletResponse resp, RedirectAttributes redirectAttributes) {
		ReturnObject returnObject = new ReturnObject();

		try {

			ExcelHelper excelHelper = new ExcelHelper(servletContext.getRealPath("/WEB-INF/exportlist"), "Bill_QuickPrint.xlsx");
			Workbook workbook = excelHelper.getWorkbook();
			if(locale.getLanguage().equalsIgnoreCase("vi")) {
			    workbook.removeSheetAt(workbook.getSheetIndex("en"));
			}else {
			    workbook.removeSheetAt(workbook.getSheetIndex("vi"));
			}
			Sheet sheet = workbook.getSheetAt(0);

			Bill bill = billService.findById(billId);
			Map<Long, BillDetail> billDetailMapKey = new LinkedHashMap<Long, BillDetail>();
			List<BillDetail> billDetaiList = new ArrayList<BillDetail>();
			int tbSize = 0;
			if (CollectionUtils.isNotEmpty(bill.getBillDetails())) {
				tbSize = bill.getBillDetails().size();
				billDetaiList.addAll(bill.getBillDetails());
				// sort trungnt
//				try {
//					BillBean bean = new BillBean();
//					billDetaiList = billDetailService.findByAllSortCategory(bill.getBillId(), bean);
//				} catch (Exception e) {
//
//				}

				List<BillDetail> childs = new ArrayList<BillDetail>();
				for (BillDetail billDetail1 : billDetaiList) {
					long billDetailId = billDetail1.getBillDetailId();
					Long parentId = billDetail1.getParentBillDetail();
					BillDetail billDetail2 = billDetailMapKey.get(billDetailId);
					if (billDetail2 == null && parentId == null) {
						billDetailMapKey.put(billDetailId, billDetail1);
					} else if (parentId != null) {
						childs.add(billDetail1);
					}
				}

				for (BillDetail child : childs) {
					long parentId = child.getParentBillDetail();
					billDetailMapKey.get(parentId).addChild(child);
				}

			}

			excelHelper.fillCell(sheet, "C4", bill.getProjectId().getStock().getName());
			excelHelper.fillCell(sheet, "I4", bill.getProjectId().getName());
			excelHelper.fillCell(sheet, "C5", bill.getBillPeriod().substring(4, 6) + "/" + bill.getBillPeriod().substring(0, 4));

			String colSTT = "A";
			String colEqptCode = "B";
			String colEqptName = "C";
			String colEqptUnit = "D";
			String colPriceOneTime = "E";
			String colInvoiceNo = "F";
			String colFromDate = "G";
			String colToDate = "H";
			String colQuantity = "I";
			String colNumDaysUsed = "J";
			String colQuantityAdjust = "K";
			String colNumDaysAdjust = "L";
			String colPrice = "M";
			String colAmount = "N";
			String colDesc = "O";
			
			CellReference landMark = new CellReference("A8");
			int startRow = landMark.getRow();
			int stt = 1;
			if (MapUtils.isNotEmpty(billDetailMapKey)) {

				for (int i = 1; i < (tbSize); i++) {
					sheet.createRow(startRow + 1);
					excelHelper.copyRow(sheet, startRow, startRow + 1);
				}
				ExcelHelper.removeRow(sheet, landMark.getRow());

				Set<Long> billDetailIds = new HashSet<Long>(billDetailMapKey.keySet());

				for (Long billDetailId : billDetailIds) {
					BillDetail billDetail = billDetailMapKey.get(billDetailId);

					startRow++;
					excelHelper.fillCell(sheet, colSTT + startRow, stt++);
					excelHelper.fillCell(sheet, colEqptCode + startRow, billDetail.getEquipmentId().getEquipmentCode());
					excelHelper.fillCell(sheet, colEqptName + startRow, billDetail.getEquipmentId().getName());
					excelHelper.fillCell(sheet, colEqptUnit + startRow, billDetail.getEquipmentId().getUnitName());
					excelHelper.fillCell(sheet, colPriceOneTime + startRow, (BooleanUtils.isTrue(billDetail.getPriceOneTime()) ? "1" : "0"));
					excelHelper.fillCell(sheet, colInvoiceNo + startRow, billDetail.getInvoiceNo());
					excelHelper.fillCell(sheet, colFromDate + startRow, billDetail.getFromDate());
					excelHelper.fillCell(sheet, colToDate + startRow, billDetail.getToDate());
					excelHelper.fillCell(sheet, colQuantity + startRow, billDetail.getQuantity());
					excelHelper.fillCell(sheet, colNumDaysUsed + startRow, billDetail.getNumDaysUsed());
					excelHelper.fillCell(sheet, colQuantityAdjust + startRow, billDetail.getQuantityAdjust());
					excelHelper.fillCell(sheet, colNumDaysAdjust + startRow, billDetail.getNumDaysAdjust());
					excelHelper.fillCell(sheet, colPrice + startRow, billDetail.getPrice());
					excelHelper.fillCell(sheet, colAmount + startRow, billDetail.getAmount());
					excelHelper.fillCell(sheet, colDesc + startRow, billDetail.getDescription());

					if (CollectionUtils.isNotEmpty(billDetail.getChilds())) {
						for (BillDetail bdChild : billDetail.getChilds()) {
							startRow++;
							excelHelper.fillCell(sheet, colQuantityAdjust + startRow, bdChild.getQuantityAdjust());
							excelHelper.fillCell(sheet, colNumDaysAdjust + startRow, bdChild.getNumDaysAdjust());
							excelHelper.fillCell(sheet, colPrice + startRow, bdChild.getPrice());
							excelHelper.fillCell(sheet, colAmount + startRow, bdChild.getAmount());
							excelHelper.fillCell(sheet, colDesc + startRow, bdChild.getDescription());
						}
					}

				}

				startRow++;
				excelHelper.fillCell(sheet, colAmount + startRow, bill.getTotalAmount());

			}

			String outFileName = "Bill_QuickPrint" + new Date().toString() + ".xlsx";

			resp.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=UTF-8");
			resp.setHeader("Content-Disposition", "attachment; filename=\"" + outFileName + "\"");
			workbook.write(resp.getOutputStream());
			resp.flushBuffer();

		} catch (Exception e) {
			logger.debug("##quickPrint##", e);
			returnObject.setMessage(e.getMessage());
			returnObject.setStatus(ReturnObject.ERROR);
		}

		return returnObject;
	}

	@RequestMapping(value = "/ReportBillAccumulated", method = { RequestMethod.POST })
	public void doReportBillAccumulatedPost(@ModelAttribute(value = "bean") BillBean bean, Model model, Locale locale, HttpServletRequest req,
			HttpServletResponse resp, RedirectAttributes redirectAttributes) {

		try {

			ExcelHelper excelHelper = new ExcelHelper(servletContext.getRealPath("/WEB-INF/exportlist"), "ReportBillAccumulated.xlsx");
			Workbook workbook = excelHelper.getWorkbook();
			if(locale.getLanguage().equalsIgnoreCase("vi")) {
			    workbook.removeSheetAt(workbook.getSheetIndex("en"));
			}else {
			    workbook.removeSheetAt(workbook.getSheetIndex("vi"));
			}
			Sheet sheet = workbook.getSheetAt(0);
			// Calendar cal = Calendar.getInstance();
			// cal.setTime(bean.getBillPeriod());
			Calendar calFrom = Calendar.getInstance();
			calFrom.setTime(bean.getBillFrom());
			Calendar calTo = Calendar.getInstance();
			calTo.setTime(bean.getBillTo());

			// List<Object[]> objects = billService.getReportProjectBillAccumulated(bean.getProjectId(),
			// cal.get(Calendar.YEAR), cal.get(Calendar.MONTH) + 1);
			List<Object[]> objects = billService.getReportProjectBillAccumulated(bean.getProjectId(), calFrom.get(Calendar.YEAR),
					calFrom.get(Calendar.MONTH) + 1, calTo.get(Calendar.YEAR), calTo.get(Calendar.MONTH) + 1);

			// CongDT, [2017-02-06 11:59:32.742]
			Boolean isIncludeInternal = bean.getIsIncludeInternal();
			Boolean isIncludeRent = bean.getIsIncludeRent();
			if (BooleanUtils.isTrue(isIncludeInternal) && BooleanUtils.isTrue(isIncludeRent)) {
				// Keep all
			} else if (BooleanUtils.isTrue(isIncludeInternal)) {

				for (Iterator<Object[]> iterator = objects.iterator(); iterator.hasNext();) {
					Object[] object = iterator.next();
					String categoryName = (String) object[0];
					if (StringUtils.contains(categoryName.toUpperCase(), "THUÊ NGOÀI") == true) {
						iterator.remove();
					}
				}

			} else if (BooleanUtils.isTrue(isIncludeRent)) {

				for (Iterator<Object[]> iterator = objects.iterator(); iterator.hasNext();) {
					Object[] object = iterator.next();
					String categoryName = (String) object[0];
					if (StringUtils.contains(categoryName.toUpperCase(), "THUÊ NGOÀI") == false) {
						iterator.remove();
					}
				}

			} else {
				objects = Collections.emptyList();
			}

			// CellStyle cellStyle = ExcelHelper.getCellStyle(sheet, "D7");
			CellStyle cellStyleDate = ExcelHelper.getCellStyle(sheet, "D7");

			CellStyle cellStyleB7 = ExcelHelper.getCellStyle(sheet, "B7");
			CellStyle cellStyleC7 = ExcelHelper.getCellStyle(sheet, "C7");
			CellStyle cellStyleD7 = ExcelHelper.getCellStyle(sheet, "D7");
			CellStyle cellStyleE7 = ExcelHelper.getCellStyle(sheet, "E7");
			CellStyle cellStyleF7 = ExcelHelper.getCellStyle(sheet, "F7");
			CellStyle cellStyleG7 = ExcelHelper.getCellStyle(sheet, "G7");

			int count = 1;

			if (objects.size() > 1) {
				if (CollectionUtils.isNotEmpty(objects)) {
					sheet.shiftRows(7, sheet.getLastRowNum(), objects.size() - 1);
				}
			}

			DateTime fromDate = new DateTime(bean.getBillFrom());
			fromDate = fromDate.plusMonths(-1);
			fromDate = fromDate.withDayOfMonth(SystemConfig.BILL_PERIOD_DATE_START);

			DateTime toDate = new DateTime(bean.getBillTo());
			toDate = toDate.withDayOfMonth(SystemConfig.BILL_PERIOD_DATE_END);

			SimpleDateFormat sdf_VN = new SimpleDateFormat("dd/MM/yyyy");
			String titleDateString1 = String.format("Từ %s đến %s", sdf_VN.format(fromDate.toDate()), sdf_VN.format(toDate.toDate()));
			String titleDateString2 = String.format("Từ %s \nđến %s", sdf_VN.format(fromDate.toDate()), sdf_VN.format(toDate.toDate()));

			excelHelper.fillCell(sheet, "B3", titleDateString1);

			Project project = projectService.findById(bean.getProjectId());
			excelHelper.fillCell(sheet, "C4", "PM " + project.getName().toUpperCase());
			excelHelper.fillCell(sheet, "D6", titleDateString2);

			SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMM");
			String outFileName = "BillLuyKe_" + project.getProjectCode() + "_" + dateFormat.format(bean.getBillFrom()) + "_"
					+ dateFormat.format(bean.getBillTo()) + "_" + System.nanoTime() + ".xlsx";
			String newFilePdf = "BillLuyKe_" + project.getProjectCode() + "_" + dateFormat.format(bean.getBillFrom()) + "_"
					+ Utils.convertDate2Str(new Date(), "dd-MM-yyyy") + ".pdf";

			// excelHelper.fillCell(sheet, "E6", "Lũy kế\n từ đầu " + new
			// SimpleDateFormat("yyyy").format(bean.getBillPeriod()));
			excelHelper.fillCell(sheet, "E6", "Lũy kế\n từ đầu " + new SimpleDateFormat("yyyy").format(bean.getBillFrom()));

			BigDecimal amountMonth = BigDecimal.ZERO;
			BigDecimal amountInYear = BigDecimal.ZERO;
			BigDecimal amountYearDate = BigDecimal.ZERO;
			BigDecimal amountEstimate = BigDecimal.ZERO;

			count += 5;

			int stt = 1;
			for (Object[] item : objects) {

				Row row = sheet.createRow(count);
				createCell(workbook, row, (stt++), 1, cellStyleB7, cellStyleDate);
				createCell2(workbook, row, item[0], 2, cellStyleC7, cellStyleDate);
				createCell2(workbook, row, item[1], 3, cellStyleD7, cellStyleDate);
				createCell2(workbook, row, item[2], 4, cellStyleE7, cellStyleDate);
				createCell2(workbook, row, item[3], 5, cellStyleF7, cellStyleDate);
				createCell2(workbook, row, item[4], 6, cellStyleG7, cellStyleDate);

				amountMonth = amountMonth.add(new BigDecimal(item[1] == null ? "0" : item[1].toString()));
				amountInYear = amountInYear.add(new BigDecimal(item[2] == null ? "0" : item[2].toString()));
				amountYearDate = amountYearDate.add(new BigDecimal(item[3] == null ? "0" : item[3].toString()));
				amountEstimate = amountEstimate.add(new BigDecimal(item[4] == null ? "0" : item[4].toString()));

				count++;
			}

			sheet.getRow(count++).setHeightInPoints(30);

			excelHelper.fillCell(sheet, "D" + (count), Double.valueOf(amountMonth.toString()));
			excelHelper.fillCell(sheet, "E" + (count), Double.valueOf(amountInYear.toString()));
			excelHelper.fillCell(sheet, "F" + (count), Double.valueOf(amountYearDate.toString()));
			excelHelper.fillCell(sheet, "G" + (count), Double.valueOf(amountEstimate.toString()));
			
			if (StringUtils.equals(bean.getAction(), "export_excel")) {
				Utils.createFile(workbook, outFileName, req, resp);
			} else if (StringUtils.equals(bean.getAction(), "export_pdf")) {
				try (ServletOutputStream servletOutputStream = resp.getOutputStream()) {
					String tempDir = System.getProperty("java.io.tmpdir");
					String pathToFileTmp = tempDir + System.getProperty("file.separator") + UUID.randomUUID().toString() + ".xlsx";
					File fileExcel = new File(pathToFileTmp);
					FileOutputStream outFile = new FileOutputStream(fileExcel);
					workbook.write(outFile);
					outFile.close();

					// Convert Workbook To Pdf
					pathToFileTmp = Utils.converterToPDF(pathToFileTmp);
					//File filePdf = new File(pathToFileTmp);
					byte[] ret = FileUtils.readFileToByteArray(filePdf);
					fileExcel.delete();
					filePdf.delete();
					String outFileNameTmp = URLEncoder.encode(newFilePdf, "UTF-8");
					resp.setContentType("application/pdf");
					resp.addHeader("Content-Disposition", "attachment; filename=" + outFileNameTmp);
					IOUtils.write(ret, servletOutputStream);
					resp.getOutputStream().flush();
				} catch (Exception e) {
					throw e;
				}
			}
			
			/*
			 * resp.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=UTF-8");
			 * resp.setHeader("Content-Disposition", "attachment; filename=\"" + outFileName + "\"");
			 * workbook.write(resp.getOutputStream()); resp.flushBuffer();
			 */

		} catch (Exception e) {
			logger.debug("##ReportBillAccumulated##", e);
			bean.addMessage(Message.ERROR, String.valueOf(e.getMessage()));
		}
	}

	@RequestMapping(value = "/preview", method = RequestMethod.GET, produces = "text/plain;charset=UTF-8")
	@ResponseBody
	public String preview(@ModelAttribute(value = "bean") BillBean bean, HttpServletRequest request, HttpServletResponse resp) throws Exception {

		StringBuffer htmlOutput = null;
		try {

			ExcelHelper excelHelper = new ExcelHelper(servletContext.getRealPath("/WEB-INF/exportlist"), "ReportBillAccumulated.xlsx");
			Workbook workbook = excelHelper.getWorkbook();
			if(Utils.getLocale().getLanguage().equalsIgnoreCase("vi")) {
			    workbook.removeSheetAt(workbook.getSheetIndex("en"));
			}else {
			    workbook.removeSheetAt(workbook.getSheetIndex("vi"));
			}
			Sheet sheet = workbook.getSheetAt(0);

			// Calendar cal = Calendar.getInstance();
			// cal.setTime(bean.getBillPeriod());
			Calendar calFrom = Calendar.getInstance();
			calFrom.setTime(bean.getBillFrom());
			Calendar calTo = Calendar.getInstance();
			calTo.setTime(bean.getBillTo());

			// List<Object[]> objects = billService.getReportProjectBillAccumulated(bean.getProjectId(),
			// cal.get(Calendar.YEAR), cal.get(Calendar.MONTH) + 1);
			List<Object[]> objects = billService.getReportProjectBillAccumulated(bean.getProjectId(), calFrom.get(Calendar.YEAR),
					calFrom.get(Calendar.MONTH) + 1, calTo.get(Calendar.YEAR), calTo.get(Calendar.MONTH) + 1);

			// CongDT, [2017-02-06 11:59:32.742]
			Boolean isIncludeInternal = bean.getIsIncludeInternal();
			Boolean isIncludeRent = bean.getIsIncludeRent();
			if (BooleanUtils.isTrue(isIncludeInternal) && BooleanUtils.isTrue(isIncludeRent)) {
				// Keep all
			} else if (BooleanUtils.isTrue(isIncludeInternal)) {

				for (Iterator<Object[]> iterator = objects.iterator(); iterator.hasNext();) {
					Object[] object = iterator.next();
					String categoryName = (String) object[0];
					if (StringUtils.contains(categoryName.toUpperCase(), "THUÊ NGOÀI") == true) {
						iterator.remove();
					}
				}

			} else if (BooleanUtils.isTrue(isIncludeRent)) {

				for (Iterator<Object[]> iterator = objects.iterator(); iterator.hasNext();) {
					Object[] object = iterator.next();
					String categoryName = (String) object[0];
					if (StringUtils.contains(categoryName.toUpperCase(), "THUÊ NGOÀI") == false) {
						iterator.remove();
					}
				}

			} else {
				objects = Collections.emptyList();
			}

			CellStyle cellStyle = ExcelHelper.getCellStyle(sheet, "D7");
			CellStyle cellStyleDate = ExcelHelper.getCellStyle(sheet, "D7");

			int count = 1;

			if (objects.size() > 1) {
				if (CollectionUtils.isNotEmpty(objects)) {
					sheet.shiftRows(7, 18, objects.size() - 1);
				}
			}

			DateTime fromDate = new DateTime(bean.getBillFrom());
			fromDate = fromDate.plusMonths(-1);
			fromDate = fromDate.withDayOfMonth(SystemConfig.BILL_PERIOD_DATE_START);

			DateTime toDate = new DateTime(bean.getBillTo());
			toDate = toDate.withDayOfMonth(SystemConfig.BILL_PERIOD_DATE_END);

			SimpleDateFormat sdf_VN = new SimpleDateFormat("dd/MM/yyyy");
			String titleDateString1 = String.format("Từ %s đến %s", sdf_VN.format(fromDate.toDate()), sdf_VN.format(toDate.toDate()));
			String titleDateString2 = String.format("Từ %s \nđến %s", sdf_VN.format(fromDate.toDate()), sdf_VN.format(toDate.toDate()));

			excelHelper.fillCell(sheet, "B3", titleDateString1);

			Project project = projectService.findById(bean.getProjectId());
			excelHelper.fillCell(sheet, "C4", "PM " + project.getName().toUpperCase());
			excelHelper.fillCell(sheet, "D6", titleDateString2);

			// excelHelper.fillCell(sheet, "E6", "Lũy kế\n từ đầu " + new
			// SimpleDateFormat("yyyy").format(bean.getBillPeriod()));
			excelHelper.fillCell(sheet, "E6", "Lũy kế\n từ đầu " + new SimpleDateFormat("yyyy").format(bean.getBillFrom()));

			BigDecimal amountMonth = BigDecimal.ZERO;
			BigDecimal amountInYear = BigDecimal.ZERO;
			BigDecimal amountYearDate = BigDecimal.ZERO;
			BigDecimal amountEstimate = BigDecimal.ZERO;

			for (Object[] item : objects) {
				Row row = sheet.createRow(count + 5);
				createCell(workbook, row, count, 1, cellStyle, cellStyleDate);

				amountMonth = amountMonth.add(new BigDecimal(item[1] == null ? "0" : item[1].toString()));
				amountInYear = amountInYear.add(new BigDecimal(item[2] == null ? "0" : item[2].toString()));
				amountYearDate = amountYearDate.add(new BigDecimal(item[3] == null ? "0" : item[3].toString()));
				amountEstimate = amountEstimate.add(new BigDecimal(item[4] == null ? "0" : item[4].toString()));

				for (int j = 0; j < item.length; j++) {
					createCell2(workbook, row, item[j], j + 2, cellStyle, cellStyleDate);
				}
				count++;
			}

			excelHelper.fillCell(sheet, "D" + (count + 6), Double.valueOf(amountMonth.toString()));
			excelHelper.fillCell(sheet, "E" + (count + 6), Double.valueOf(amountInYear.toString()));
			excelHelper.fillCell(sheet, "F" + (count + 6), Double.valueOf(amountYearDate.toString()));
			excelHelper.fillCell(sheet, "G" + (count + 6), Double.valueOf(amountEstimate.toString()));

			htmlOutput = new StringBuffer();
			PoiToHtmlConverter toHtml = PoiToHtmlConverter.create(workbook, htmlOutput);
			toHtml.setCompleteHTML(true);
			toHtml.printPage();
			/*
			 * FileOutputStream outFile = new FileOutputStream(new File(pathToFileTmp)); workbook.write(outFile);
			 * outFile.close();
			 * 
			 * ExcelUtils excelUtils = new ExcelUtils();
			 * 
			 * String extendFile = getExtendFile(pathToFileTmp); if (extendFile.equals("xls") ||
			 * extendFile.equals("xlsx")) { excelUtils.fitPrintRangeWidthAndHeight(pathToFileTmp, extendFile); }
			 * 
			 * pathToFileTmp = Utils.converterToPDF(pathToFileTmp);
			 */

		} catch (Exception e) {
			logger.debug("##ReportBillAccumulated##", e);
			bean.addMessage(Message.ERROR, String.valueOf(e.getMessage()));
		}
		return String.valueOf(htmlOutput);
	}

	private void createCell(Workbook workbook, Row row, Object object, Integer column, CellStyle cellStyle, CellStyle cellStyleDate) {
		if (object != null) {
			CreationHelper createHelper = workbook.getCreationHelper();
			Cell cell = row.createCell(column);
			;
			if (object instanceof Double) {
				// format cell
				cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)"));
				cell.setCellValue((Double) object);
				cell.setCellStyle(cellStyle);
			}
			if (object instanceof BigDecimal) {
				// format cell
				cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)"));
				cell.setCellValue((((BigDecimal) object).doubleValue()));
				cell.setCellStyle(cellStyle);
			}
			if (object instanceof Long) {
				// format cell
				cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)"));
				cell.setCellValue((Long) object);
				cell.setCellStyle(cellStyle);
			}
			if (object instanceof Integer) {
				cell = row.createCell(column, CellType.NUMERIC);
				cell.setCellValue((Integer) object);
				cell.setCellStyle(cellStyle);
			}
			if (object instanceof String) {
				cell = row.createCell(column, CellType.STRING);
				cell.setCellValue(object.toString());
				cell.setCellStyle(cellStyle);
			}
			if (object instanceof Date) {
				cell.setCellValue((Date) object);
				cell.setCellStyle(cellStyleDate);
			}
		} else {

			row.createCell(column, CellType.STRING).setCellValue("");
			row.getCell(column).setCellStyle(cellStyle);
		}
	}

	private void createCell2(Workbook workbook, Row row, Object object, Integer column, CellStyle cellStyle, CellStyle cellStyleDate) {
		if (object != null) {
			Cell cell = row.createCell(column);
			if (object instanceof Double) {
				// format cell
				cell.setCellValue((Double) object);
				cell.setCellStyle(cellStyle);
			}
			if (object instanceof BigDecimal) {
				// format cell
				cell.setCellValue((((BigDecimal) object).doubleValue()));
				cell.setCellStyle(cellStyle);
			}
			if (object instanceof Long) {
				// format cell
				cell.setCellValue((Long) object);
				cell.setCellStyle(cellStyle);
			}
			if (object instanceof Integer) {
				cell = row.createCell(column, CellType.NUMERIC);
				cell.setCellValue((Integer) object);
				cell.setCellStyle(cellStyle);
			}
			if (object instanceof String) {
				cell = row.createCell(column, CellType.STRING);
				cell.setCellValue(object.toString());
				cell.setCellStyle(cellStyle);
			}
			if (object instanceof Date) {
				cell.setCellValue((Date) object);
				cell.setCellStyle(cellStyleDate);
			}
		} else {

			row.createCell(column, CellType.STRING).setCellValue("");
			row.getCell(column).setCellStyle(cellStyle);
		}
	}

	@RequestMapping(value = "/edit", method = { RequestMethod.GET, RequestMethod.POST })
	public String doEditBill(@ModelAttribute(value = "bean") BillBean bean, Model model, Locale locale, HttpServletRequest req) {

		boolean isAllowEdit = false;
		boolean isAllowSubmit = false;
		boolean isShowApproval = false;
		boolean isAllowApproval = false;

		try {
			String backLink = null;
			if (req.isUserInRole("R019")) {
				backLink = "listStock";
			} else if (req.isUserInRole("R020")) {
				backLink = "listStock";
			} else if (req.isUserInRole("R021")) {
				backLink = "listProject";
			} else {
				throw new Exception(getMsg("Bill.msg.AccessNotPermitted"));
			}
			model.addAttribute("backLink", backLink);

			SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMM");
			Bill bill = null;
			if (bean.getBillId() != null) {
				bill = billService.findById(bean.getBillId());
			} else if (bean.getProjectId() != null && bean.getBillPeriod() != null) {
				bill = billService.findByProjectAndPeriod(bean.getProjectId(), dateFormat.format(bean.getBillPeriod()));
			} else {
				throw new Exception(getMsg("msg.data.not.found"));
			}

			if (bill == null) {
				throw new Exception(getMsg("msg.data.not.found"));
			}

			bean.setBillPeriod(dateFormat.parse(bill.getBillPeriod()));
			bean.setEntity(bill);

			if (bill.getStatus() != SystemConfig.STATUS_BILL_DRAFT) {
				Account accCreated = accountService.findByAccountName(bill.getCreatedBy());
				model.addAttribute("CreatedByName", accCreated.getFullName());
			}

			// Kiểm tra trạng thái bill
			if (bill.getStatus() == SystemConfig.STATUS_BILL_DRAFT || bill.getStatus() == SystemConfig.STATUS_BILL_RETURNED) {
				// Bill phải ở trạng thái nháp thì mới được submit
				isAllowEdit = true;
				isAllowSubmit = true;
				// User phải thuộc kho thì mới có quyền save & submit
				if (userProfile.getAccount().getStock() == null || userProfile.getAccount().getStock().getStockId() == null) {
					isAllowEdit = false;
				} else if (bill.getProjectId().getStock().getStockId().equals(userProfile.getAccount().getStock().getStockId()) == false) {
					isAllowEdit = false;
				}

				if (BooleanUtils.isNotTrue(req.isUserInRole("R020"))) {
					isAllowEdit = false;
				}
			}

			// -------Bill Detail Start ------ //
			Map<Long, BillDetail> billDetailMapKey = new LinkedHashMap<>();
			List<BillDetail> billDetaiList = new ArrayList<BillDetail>();
			List<Long> billDetailIds = new ArrayList<>();
			// if (CollectionUtils.isNotEmpty(bill.getBillDetails())) {
			// billDetaiList.addAll(bill.getBillDetails());
			// billDetaiList = billDetailService.findByStockTrackingId(stockTrackingId)
			// sort trungnt
			try {
				billDetaiList = billDetailService.findByAllSortCategory(bill.getBillId(), bean);
			} catch (Exception e) {
			}
			List<BillDetail> childs = new ArrayList<BillDetail>();
			for (BillDetail billDetail1 : billDetaiList) {
				long billDetailId = billDetail1.getBillDetailId();
				Long parentId = billDetail1.getParentBillDetail();
				BillDetail billDetail2 = billDetailMapKey.get(billDetailId);
				if (billDetail2 == null && parentId == null) {
					billDetailMapKey.put(billDetailId, billDetail1);
					billDetailIds.add(billDetailId);
				} else if (parentId != null) {
					childs.add(billDetail1);
				}
			}

			for (BillDetail child : childs) {
				long parentId = child.getParentBillDetail();
				billDetailMapKey.get(parentId).addChild(child);
			}

			// }
			model.addAttribute("billDetailMapKey", billDetailMapKey);

			model.addAttribute("billDetailIds", billDetailIds);
			// -------Bill Detail End ------ //

			List<HistoryApprove> historyApproves = historyApproveService.findBySourceAndRef(bill.getBillId(), SystemConfig.SOURCE_BILL);
			model.addAttribute("historyApproves", historyApproves);

			if (CollectionUtils.isNotEmpty(historyApproves)) {
				isShowApproval = true;
			}

			Calendar calendarPrevPeriod = Calendar.getInstance();
			calendarPrevPeriod.setTime(bean.getBillPeriod());
			calendarPrevPeriod.add(Calendar.MONTH, -1);
			String prevPeriod = new SimpleDateFormat("MM/yyyy").format(calendarPrevPeriod.getTime());
			if (bill.getAmountAccumulatedUntilPreviousPeriod() == null) {
				bill.setAmountAccumulatedUntilPreviousPeriod(BigDecimal.ZERO);
			}
			String amountAccumulated = new DecimalFormat("###,##0").format(bill.getAmountAccumulatedUntilPreviousPeriod());
			model.addAttribute("prevPeriod", prevPeriod);
			model.addAttribute("amountAccumulated", amountAccumulated);

			if (StringUtils.equals(userProfile.getAccount().getUsername(), bill.getPendingAt())) {
				isAllowApproval = true;
			}

			// Trạng thái xét duyệt
			model.addAttribute("BillApproveStatus", systemConfig.getHistoryApprovalStatusMsg(locale));

			// Lọc công trường theo stock
			List<Project> projects = projectService.findByStockId(bill.getProjectId().getStock().getStockId());
			model.addAttribute("projects", projects);

			// list process
			List<ProcessStep> fullProcessSteps = new LinkedList<ProcessStep>();
			setApproverListPopup(bill, model, fullProcessSteps);

			// Next Approval
			String nextApprover = null;
			if (StringUtils.equalsIgnoreCase(bill.getPendingAt(), userProfile.getAccount().getUsername())) {

				long processId1 = bill.getProcessId();
				int levelProcess1 = bill.getLevelProcess();
				int stepNo1 = bill.getStepNo();
				boolean isBreak = false;

				for (ProcessStep processStep : fullProcessSteps) {
					long processId2 = processStep.getProcessId();
					int levelProcess2 = processStep.getLevelProcess();
					int stepNo2 = processStep.getStepNo();

					if (isBreak == true) {
						nextApprover = processStep.getAccount().getFullName();
						break;
					} else if (processId1 == processId2 && levelProcess1 == levelProcess2 && stepNo1 == stepNo2) {
						isBreak = true;
					}
				}
			} else {
				nextApprover = bill.getPendingAtFullName();
			}
			model.addAttribute("nextApproval", nextApprover);

			List<Stock> stocks = stockService.findAll();
			model.addAttribute("stocks", stocks);
		} catch (Exception e) {
			logger.debug("##edit##", e);
			isAllowEdit = false;
			isAllowSubmit = false;
			isShowApproval = false;
			isAllowApproval = false;
			bean.addMessage(Message.ERROR, String.valueOf(e.getMessage()));
		}

		model.addAttribute("isAllowEdit", isAllowEdit);
		model.addAttribute("isAllowSubmit", isAllowSubmit);
		model.addAttribute("isShowApproval", isShowApproval);
		model.addAttribute("isAllowApproval", isAllowApproval);

		return "Bill.edit";
	}

	@RequestMapping(value = "/update", method = { RequestMethod.POST })
	@ResponseBody
	public Object doUpdateBill(@ModelAttribute(value = "bean") @Valid BillBean bean, BindingResult bindingResult, Model model, Locale locale) {

		ReturnObject retObj = new ReturnObject();
		try {

			doSaveDraft(bean);

			retObj.setStatus(ReturnObject.SUCCESS);
			retObj.setMessage(getMsg("msg.save.success"));

		} catch (Exception e) {
			logger.debug("##save##", e);
			retObj.setStatus(ReturnObject.ERROR);
			retObj.setMessage(e.getMessage());
		}
		return retObj;

	}

	private void doSaveDraft(BillBean bean) throws Exception {
		Bill bill = bean.getEntity();		
		bill.setUpdatedBy(userProfile.getAccount().getUsername());
		bill.setUpdatedDate(new Date());
		// logger.debug("##BillDetail##" +
		// Utils.writeValueAsString(bean.getBillDetails()));

		List<BillDetail> billDetails = bean.getBillDetails();
		
		List<BillDetail> billDetailAddNewList = new ArrayList<BillDetail>();
		List<BillDetail> billDetailUpdateList = new ArrayList<BillDetail>();

		for (BillDetail billDetail : billDetails) {
			billDetail.setBillId(bill);
			if (billDetail.getBillDetailId() == null && billDetail.getParentBillDetail() != null) {
				billDetailAddNewList.add(billDetail);
			} else if (billDetail.getBillDetailId() != null) {
				billDetailUpdateList.add(billDetail);
			}
		}

		bean.setBillDetailUpdateList(billDetailUpdateList);
		bean.setBillDetailAddNewList(billDetailAddNewList);

		bill.setBillDetails(new HashSet<BillDetail>(bean.getBillDetails()));

		bill.setUpdatedBy(userProfile.getAccount().getUsername());
		bill.setUpdatedDate(new Date());

		billService.update(bean);
	}
	
	//xóa 1 record trong chi phí sử dụng sản phẩm
		@RequestMapping(value = "/delete", method = RequestMethod.POST)
		@ResponseBody
		public Boolean deleteRow(@RequestParam(value = "billDetailId", required = false) Long billDetailId) throws Exception {
			
			int count = billDetailService.deleteByBillDetailId(billDetailId);
			if(count > 0) 
				return true;
			else
				return false;
		}

	@RequestMapping(value = "findProjectByStock", method = RequestMethod.GET)
	@ResponseBody
	public Object doFindProjectByStock(@RequestParam(value = "stockId") Long stockId, Model model, Locale locale) {
		ReturnObject returnObject = new ReturnObject();

		try {
			if (stockId == null) {
				throw new Exception(msgSrc.getMessage("msg.stockid.null", null, locale));
			}

			List<Project> projects = projectService.findByStockId(stockId);

			returnObject.setRetObj(projects);
			returnObject.setStatus(ReturnObject.SUCCESS);
			returnObject.setMessage("Find Done");

		} catch (Exception e) {
			returnObject.setStatus(ReturnObject.ERROR);
			returnObject.setMessage(String.valueOf(e.getMessage()));
		}

		return returnObject;
	}

	@RequestMapping(value = "run", method = RequestMethod.GET)
	@ResponseBody
	public Object doBillTaskManually(@RequestParam(value = "month") Integer month, @RequestParam(value = "projectid") String projectid) {

		try {
			Calendar calendar1 = Calendar.getInstance();
			SimpleDateFormat sdfYYYYMM = new SimpleDateFormat("yyyyMM");
			calendar1.set(Calendar.MONTH, month - 1);
			String currentPeriod = sdfYYYYMM.format(calendar1.getTime());

			calendar1.set(Calendar.DATE, SystemConfig.BILL_PERIOD_DATE_END);

			Calendar calendar2 = Calendar.getInstance();
			calendar2.set(Calendar.MONTH, month);
			calendar2.set(Calendar.DATE, SystemConfig.BILL_PERIOD_DATE_START);
			String nextPeriod = sdfYYYYMM.format(calendar2.getTime());

			long diffDaysRoot = (calendar2.getTime().getTime() - calendar1.getTime().getTime()) / (24 * 60 * 60 * 1000) + 1;
			long diffDays = diffDaysRoot;
			List<Map<String, Object>> currentBills = new ArrayList<>();
			if (projectid.equals("") || projectid == null) {
				currentBills = billService.collectBillDataForNextPeriod(currentPeriod);
			} else {
				currentBills = billService.collectBillDataForNextPeriod(currentPeriod, projectid);
			}

			if (CollectionUtils.isNotEmpty(currentBills)) {
				Map<Long, Bill> billMap = new HashMap<Long, Bill>();
				for (Map<String, Object> map : currentBills) {
					long billId = (long) map.get("BillId");
					long projectId = (long) map.get("ProjectId");
					Long equipmentId = (Long) map.get("EquipmentId");
					Long equipmentCategoryId = (Long) map.get("EquipmentCategoryId");
					Double quantity = Double.parseDouble(map.get("Quantity").toString());

					Bill bill = billMap.get(billId);
					if (bill == null) {
						bill = new Bill();
						Project project = new Project();
						project.setProjectId(projectId);
						bill.setProjectId(project);
						bill.setBillPeriod(nextPeriod);
						bill.setBillRunDate(calendar1.getTime());
						bill.setStatus(0);
						bill.setCreatedDate(calendar1.getTime());
						bill.setTotalAmount(BigDecimal.ZERO);
						billMap.put(billId, bill);
					}

					if (equipmentId != null) {
						if (bill.getBillDetails() == null) {
							bill.setBillDetails(new HashSet<BillDetail>());
						}

						BillDetail billDetail = new BillDetail();
						Equipment equipment = new Equipment();
						equipment.setEquipmentId(equipmentId);
						billDetail.setEquipmentId(equipment);
						billDetail.setEquipmentCategoryId(equipmentCategoryId);
						billDetail.setQuantity(quantity);
						billDetail.setQuantityAdjust(quantity);
						billDetail.setFromDate(calendar1.getTime());
						billDetail.setToDate(calendar2.getTime());

						EquipmentPriceBean equipmentPriceBean = priceProjectSettingService.findEquipmentPrice(projectId, equipmentId, new Date());
						if (equipmentPriceBean != null && equipmentPriceBean.getPrice() != null) {
							billDetail.setPrice(equipmentPriceBean.getPrice());
							billDetail.setPriceOneTime(equipmentPriceBean.getPriceOneTimes());
							if (BooleanUtils.isTrue(equipmentPriceBean.getPriceOneTimes())) {
								diffDays = 1;
							} else {
								diffDays = diffDaysRoot;
							}
						} else {
							billDetail.setPrice(new BigDecimal(0));
						}

						billDetail.setNumDaysUsed((int) diffDays);
						billDetail.setNumDaysAdjust((int) diffDays);

						BigDecimal amount = BigDecimal.valueOf(quantity).multiply(billDetail.getPrice()).multiply(new BigDecimal(diffDays));
						billDetail.setAmount(amount);
						bill.setTotalAmount(bill.getTotalAmount().add(amount));

						billDetail.setIsOpeningBalance(true);
						billDetail.setBillId(bill);

						bill.getBillDetails().add(billDetail);

					}

				}

				billService.saveOpeningBalanceBill(new ArrayList<Bill>(billMap.values()));

			}

		} catch (Exception e) {
			logger.debug("##BillTask##", e);
		}

		return "OK run";

	}

	/**
	 * Main Run Bill CongDT 2016-12-29
	 *
	 * @param month
	 * @param projectid
	 * @return
	 * @throws Exception
	 */
	@RequestMapping(value = "/runbill", method = RequestMethod.GET)
	@ResponseBody
	@Transactional(rollbackFor = Exception.class)
	public Object runbill(@RequestParam(value = "billPreriod") String billPeriod, @RequestParam(value = "projectId", required = false) Long projectId,
			HttpServletRequest req, HttpServletResponse resp) throws Exception {

		Map<String, Object> retObj = new LinkedHashMap<>();

		Date datePeriod = new SimpleDateFormat("yyyyMM").parse(billPeriod);
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(datePeriod);
		calendar.add(Calendar.MONTH, -1);
		calendar.set(Calendar.DATE, SystemConfig.BILL_PERIOD_DATE_END);
		retObj.put("1.LastBillCutoff", String.valueOf(calendar.getTime()));

		System.out.println(calendar);

		try {
			// Lấy dữ liệu tồn đầu kỳ
			try {
				List<StockTracking> stockTrackingsstock = new ArrayList<>();
				if (projectId == null) {
					stockTrackingsstock = stockTrackingService.findAllRunBillStock(calendar);
				} else {
					stockTrackingsstock = stockTrackingService.findAllRunBillStock(calendar, projectId, null);
				}

				int i = 1;
				for (StockTracking stockTracking : stockTrackingsstock) {

					if (stockTracking.getEquipmentId() == null) {
						throw new Exception(Utils.writeValueAsString(stockTracking));
					}
					System.out.println("#R1#:" + (i++));
					billService.convertStockTrackingToRunBill(stockTracking, billPeriod, true);
				}
			} catch (Exception e) {
				throw e;
			}

			try {

				DateTime dateTimeFrom = new DateTime(datePeriod);
				dateTimeFrom = dateTimeFrom.minusMonths(1);
				dateTimeFrom = dateTimeFrom.withField(DateTimeFieldType.dayOfMonth(), SystemConfig.BILL_PERIOD_DATE_START);
				Date dateFrom = dateTimeFrom.toDate();
				retObj.put("2.NewBillFrom", String.valueOf(dateFrom));

				DateTime dateTimeTo = new DateTime(datePeriod);
				dateTimeTo = dateTimeTo.withField(DateTimeFieldType.dayOfMonth(), SystemConfig.BILL_PERIOD_DATE_END);
				Date dateTo = dateTimeTo.toDate();
				retObj.put("3.NewBillTo", String.valueOf(dateTo));

				List<StockTracking> stockTrackings = new ArrayList<>();
				if (projectId == null) {
					stockTrackings = stockTrackingService.findAllRunBill(dateFrom, dateTo);
				} else {
					stockTrackings = stockTrackingService.findAllRunBill(dateFrom, dateTo, projectId, null);
				}

				int i = 1;
				for (StockTracking stockTracking : stockTrackings) {
					System.out.println("#R2#:" + (i++));
					billService.convertStockTrackingToRunBill(stockTracking, billPeriod, false);
				}
			} catch (Exception e) {
				throw e;
			}
		} catch (Exception e) {
			retObj.put("Exception", ExceptionUtils.getStackTrace(e));
		}

		return Utils.writeValueAsString(retObj);

	}

	/**
	 * Xóa Bill
	 * 
	 * @author CongDT
	 * @since 2016-01-06
	 * @param billPeriod
	 * @param projectId
	 * @param token
	 * @return
	 * @throws Exception
	 */
	@RequestMapping(value = "deletebill", method = RequestMethod.GET)
	@ResponseBody
	public Object deletebill(@RequestParam(value = "billPeriod", required = true) String billPeriod,
			@RequestParam(value = "projectId", required = false) Long projectId, @RequestParam(value = "token", required = false) String token, Locale locale)
			throws Exception {

		if (StringUtils.isBlank(billPeriod) || StringUtils.length(billPeriod) != 6) {
			throw new Exception(msgSrc.getMessage("msg.period.invalid", null, locale) + "," + billPeriod);
		}

		try {
			new SimpleDateFormat("yyyyMM").parse(billPeriod);
		} catch (Exception e) {
			throw e;
		}

		int numberDeleted = 0;
		if (projectId == null && StringUtils.equalsIgnoreCase(token, "ALL")) {
			numberDeleted = billService.deleteBill(billPeriod);
		} else if (projectId != null) {
			numberDeleted = billService.deleteBill(billPeriod, projectId);
		} else {
			throw new Exception(msgSrc.getMessage("msg.not.process", null, locale));
		}

		return String.format("DELETE BILL DONE. billPeriod=%s, projectId=%s, numberDeleted=%d", billPeriod, projectId, numberDeleted);

	}

	@RequestMapping(value = "runBillTaskManually", method = RequestMethod.GET)
	@ResponseBody
	public Object doBillTaskManually() throws Exception {

		BillTask billTask = new BillTask();
		appContext.getAutowireCapableBeanFactory().autowireBean(billTask);
		billTask.doTask();

		return "OK";

	}

	private String getMsg(String code) {
		return msgSrc.getMessage(code, null, LocaleContextHolder.getLocale());
	}

	@RequestMapping(value = "submitBill", method = RequestMethod.POST)
	@ResponseBody
	public Object doSubmitBill(@ModelAttribute(value = "bean") @Valid BillBean bean, BindingResult bindingResult, Model model, Locale locale) {
		ReturnObject returnObject = new ReturnObject();

		try {
			doSaveDraft(bean);
			Bill bill = billService.findById(bean.getEntity().getBillId());
//			if (bill.getStatus() != SystemConfig.STATUS_BILL_DRAFT) {
//				throw new Exception(msgSrc.getMessage("msg.status.bill.draft", null, locale));
//			}

			bill.setStatus(SystemConfig.STATUS_BILL_WAITING_APPROVED);
			bill.setCreatedBy(userProfile.getAccount().getUsername());
			bill.setCreatedDate(new Date());
			bill.setUpdatedBy(userProfile.getAccount().getUsername());
			bill.setUpdatedDate(new Date());

			List<String> approveList = processStepPendingService.getApproverList(bill.getBillId(), SystemConfig.SOURCE_BILL, bill.getStepNo());

			billService.submit(bill, approveList, bean.getAction());

			returnObject.setStatus(ReturnObject.SUCCESS);
			returnObject.setMessage(getMsg("Approval.submited.succesfully"));

		} catch (Exception e) {
			logger.debug("##submitBill##", e);
			returnObject.setStatus(ReturnObject.ERROR);
			returnObject.setMessage(e.getMessage());
		}

		return returnObject;
	}

	@RequestMapping(value = "approveBill", method = RequestMethod.POST)
	@ResponseBody
	public Object doApproveBill(@RequestParam(value = "billId", required = true) long billId,
			@RequestParam(value = "action", required = false) String action, @RequestParam(value = "comment", required = false) String comment,
			Model model, Locale locale) {

		ReturnObject returnObject = new ReturnObject();

		try {
			Bill bill = billService.findById(billId);

			if (StringUtils.equals(action, "approve")) {
				bill.setStatus(SystemConfig.STATUS_BILL_WAITING_APPROVED);
				returnObject.setMessage(getMsg("Approval.approved.succesfully"));
			} else if (StringUtils.equals(action, "return")) {
				if (bill.getStatus() != SystemConfig.STATUS_BILL_WAITING_APPROVED) {
					throw new Exception(msgSrc.getMessage("msg.can.not.return", null, locale));
				}
				bill.setStatus(SystemConfig.STATUS_BILL_RETURNED);
				returnObject.setMessage(getMsg("Approval.returned.succesfully"));
			} else {
				throw new Exception(msgSrc.getMessage("msg.action.false", null, locale));
			}

			bill.setUpdatedBy(userProfile.getAccount().getUsername());
			bill.setUpdatedDate(new Date());
			bill.setComment(StringUtils.trimToNull(comment));

			List<String> approveList = processStepPendingService.getApproverList(bill.getBillId(), SystemConfig.SOURCE_BILL, bill.getStepNo());

			billService.submit(bill, approveList, action);

			returnObject.setStatus(ReturnObject.SUCCESS);

		} catch (Exception e) {
			logger.debug("##approveBill##", e);
			returnObject.setStatus(ReturnObject.ERROR);
			returnObject.setMessage(e.getMessage());
		}

		return returnObject;
	}

	private void setApproverListPopup(Bill bill, Model model, List<ProcessStep> fullProcessSteps) {

		List<ProcessStep> processFrom = new ArrayList<ProcessStep>();
		List<ProcessStep> processTo = new ArrayList<ProcessStep>();
		List<ProcessStep> processEquipment = new ArrayList<ProcessStep>();
		List<String> strings = new ArrayList<String>();
		Project project = bill.getProjectId();
		Stock stock = project.getStock();

		processFrom = processStepService.getLstProcessSteps(stock.getProcessId());
		processTo = processStepService.getLstProcessSteps(project.getProcessId());

		for (ProcessStep processStep : processFrom) {
			if (processStep.getLeader() == SystemConfig.LEAD_STOCK) {
				processStep.setAccount(stock.getAccount());
			} else if (processStep.getLeader() == SystemConfig.MONITORING_STOCK) {
				processStep.setAccount(stock.getMonitoring());
			}
			processStep.setLevelProcess(1);
		}
		for (ProcessStep processStep : processTo) {
			if (processStep.getLeader() == SystemConfig.LEAD_PROJECT) {
				Account account = new Account();
				account = accountService.findByAccountName(project.getProjectLeader());
				processStep.setAccount(account);
				strings.add(project.getProjectLeader());
			} else if (processStep.getLeader() == SystemConfig.MONITORING_PROJECT) {
				processStep.setAccount(project.getMonitoring());
				strings.add(project.getMonitoring().getUsername());
			}
			processStep.setLevelProcess(2);
		}

		List<Long> equipmentCategoryId = billDetailService.findBillEquipmentCategoryId(bill == null ? null : bill.getBillId());
		processEquipment = processStepService.findByProcessEquipment(new Date(), SystemConfig.PROCESS_STEP_EQUIPMENT, SystemConfig.SOURCE_BILL,
				bill.getProjectId().getStock().getStockId(), equipmentCategoryId);
		for (ProcessStep processStep : processEquipment) {
			processStep.setLevelProcess(3);
		}

		model.addAttribute("processFrom", processFrom);
		model.addAttribute("processTo", processTo);
		model.addAttribute("processEquipment", processEquipment);
		model.addAttribute("LEVEL_PROCESS_TRANSFER_FROM", SystemConfig.LEVEL_PROCESS_TRANSFER_FROM);
		model.addAttribute("LEVEL_PROCESS_TRANSFER_TO", SystemConfig.LEVEL_PROCESS_TRANSFER_TO);
		model.addAttribute("LEVEL_PROCESS_TRANSFER_EQUIPMENT", SystemConfig.LEVEL_PROCESS_TRANSFER_EQUIPMENT);

		fullProcessSteps.addAll(processFrom);
		fullProcessSteps.addAll(processTo);
		fullProcessSteps.addAll(processEquipment);

	}

	@RequestMapping(value = "/expportBillList", method = { RequestMethod.GET, RequestMethod.POST })
	public @ResponseBody Object expportBillList(@ModelAttribute(value = "bean") BillBean bean, Model model, Locale locale, HttpServletRequest req,
			HttpServletResponse resp, RedirectAttributes redirectAttributes) {
		ReturnObject returnObject = new ReturnObject();

		try {

			ExcelHelper excelHelper = new ExcelHelper(servletContext.getRealPath("/WEB-INF/exportlist"), "Bill_ExportList.xlsx");
			Workbook workbook = excelHelper.getWorkbook();
			if(locale.getLanguage().equalsIgnoreCase("vi")) {
			    workbook.removeSheetAt(workbook.getSheetIndex("en"));
			}else {
			    workbook.removeSheetAt(workbook.getSheetIndex("vi"));
			}
			Sheet sheet = workbook.getSheetAt(0);

			bean.setLimit(99999);
			List<Bill> bills = billService.find(bean);

			String colSTT = "A";
			String colStock = "B";
			String colProject = "C";
			String colPeriod = "D";
			String colAmount = "E";
			String colRundate = "F";
			String colDescription = "G";
			String colStatus = "H";

			CellReference lankMark = new CellReference("A6");
			Row rowTemp = sheet.getRow(lankMark.getRow());
			int colNum = rowTemp.getLastCellNum();
			int startRow = rowTemp.getRowNum();

			Map<String, CellStyle> cellStyleMap = new HashMap<>();
			for (int i = 0; i < colNum; i++) {
				cellStyleMap.put(CellReference.convertNumToColString(i), rowTemp.getCell(i).getCellStyle());
			}

			CellStyle cellStyleD7 = ExcelHelper.getCellStyle(sheet, "D7");
			CellStyle cellStyleE7 = ExcelHelper.getCellStyle(sheet, "E7");

			int stt = 1;
			if (CollectionUtils.isNotEmpty(bills)) {

				Map<Integer, String> statusTransferLinkMap = systemConfig.getStatusTransferLinkMap();

				BigDecimal sumAmount = BigDecimal.ZERO;

				for (Bill bill : bills) {
					startRow++;
					excelHelper.fillCell(sheet, colSTT + startRow, stt++, cellStyleMap.get(colSTT));
					excelHelper.fillCell(sheet, colStock + startRow, bill.getProjectId().getStock().getName(), cellStyleMap.get(colStock));
					excelHelper.fillCell(sheet, colProject + startRow, bill.getProjectId().getName(), cellStyleMap.get(colProject));
					excelHelper.fillCell(sheet, colPeriod + startRow, bill.getBillPeriod(), cellStyleMap.get(colPeriod));
					excelHelper.fillCell(sheet, colAmount + startRow, bill.getTotalAmount(), cellStyleMap.get(colAmount));
					excelHelper.fillCell(sheet, colRundate + startRow, bill.getBillRunDate(), cellStyleMap.get(colRundate));
					excelHelper.fillCell(sheet, colDescription + startRow, bill.getDescription(), cellStyleMap.get(colStatus));
					excelHelper.fillCell(sheet, colStatus + startRow, getMsg(statusTransferLinkMap.get(bill.getStatus())),
							cellStyleMap.get(colStatus));
					sumAmount = sumAmount.add(bill.getTotalAmount());
				}
				for (int i = 0; i < colNum; i++) {
					excelHelper.fillCell(sheet, startRow, i, "", cellStyleD7);
				}
				startRow++;
				excelHelper.fillCell(sheet, colPeriod + startRow, msgSrc.getMessage("msg.total", null, locale));
				excelHelper.fillCell(sheet, colAmount + startRow, sumAmount, cellStyleE7);

			}
			
			String outFileName = "Bill_ExportList" + (new Date()) + ".xlsx";
			try (ServletOutputStream servletOutputStream = resp.getOutputStream()) {
				resp.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=UTF-8");
				resp.setHeader("Content-Disposition", "attachment; filename=\"" + outFileName + "\"");
				workbook.write(servletOutputStream);
				resp.flushBuffer();
			} catch (Exception e) {
				throw e;
			}

		} catch (Exception e) {
			logger.debug("##quickPrint##", e);
			returnObject.setMessage(e.getMessage());
			returnObject.setStatus(ReturnObject.ERROR);
		}

		return returnObject;
	}
}
