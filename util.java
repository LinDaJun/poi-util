	private StringBuilder replace(StringBuilder sb, String holder, String word, Boolean isMulti) {
		int start = sb.indexOf(holder);
		if (start == -1) {
			return sb;
		}
		do {
			sb.replace(start, start + holder.length() , word);
			start = sb.indexOf(holder);
		}while ( start != -1 && isMulti);
		return sb;
	}

	/**
	 * 複製table至指定cursor
	 * @param doc
	 * @param table
	 * @param cursor
	 * @return
	 */
	private XWPFTable copyTable(XWPFDocument doc, XWPFTable table, XmlCursor cursor) {
		XWPFTable targetTable = doc.insertNewTbl(cursor);
		targetTable.getCTTbl().setTblPr(table.getCTTbl().getTblPr());
        //複製行
        for(int i = 0; i < table.getRows().size(); i++) {
            XWPFTableRow targetRow = targetTable.getRow(i);
            XWPFTableRow sourceRow = table.getRow(i);
            if(targetRow == null) {
            	targetRow = targetTable.createRow();
            }
            copyTableRow(targetRow, sourceRow);
        }
        
        return targetTable;
	}

	/**
	 * 複製table至doc的最尾處
	 * @param doc
	 * @param table
	 * @param cursor
	 * @return
	 */
	private XWPFTable copyTable(XWPFDocument doc, XWPFTable table) {
		XWPFTable targetTable = doc.createTable(); 
        //複製表格屬性
        targetTable.getCTTbl().setTblPr(table.getCTTbl().getTblPr());
        //複製行
        for(int i = 0; i < table.getRows().size(); i++) {
            XWPFTableRow targetRow = targetTable.getRow(i);
            XWPFTableRow sourceRow = table.getRow(i);
            if(targetRow == null) {
            	targetRow = targetTable.createRow();
            }
            copyTableRow(targetRow, sourceRow);
        }
        
        return targetTable;
	}
	
	/**
	 * 複製sourceRow至targetRow
	 * @param targetRow
	 * @param sourceRow
	 */
	private void copyTableRow(XWPFTableRow targetRow, XWPFTableRow sourceRow) {
	    //複製樣式
	    if(sourceRow != null) {
	        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
	    }
	    //複製單元格
	    for(int i = 0; i < sourceRow.getTableCells().size(); i++) {
	        XWPFTableCell tCell = targetRow.getCell(i);
	        XWPFTableCell sCell = sourceRow.getCell(i);
	        if(tCell == null) {
	            tCell = targetRow.addNewTableCell();
	        }
	        copyTableCell(tCell, sCell);
	    }
	}

	/**
	 * 複製sourceCell至targetCell
	 * @param targetCell
	 * @param sourceCell
	 */
	private void copyTableCell(XWPFTableCell targetCell, XWPFTableCell sourceCell) {
        //表格屬性
        if(sourceCell.getCTTc() != null) {
            targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
        }
        //刪除段落
        for(int pos = 0; pos < targetCell.getParagraphs().size(); pos++) {
            targetCell.removeParagraph(pos);
        }
        //新增段落
        for(XWPFParagraph sourceParag : sourceCell.getParagraphs()) {
            XWPFParagraph targetParag = targetCell.addParagraph();
            copyParagraph(targetParag, sourceParag);
        }
    }
	/**
	 * 複製sourceParag至targetParag
	 * @param targetParag
	 * @param sourceParag
	 */
	private void copyParagraph(XWPFParagraph targetParag, XWPFParagraph sourceParag) {
        targetParag.getCTP().setPPr(sourceParag.getCTP().getPPr());    //設定段落樣式
        //移除所有的run
        for(int pos = targetParag.getRuns().size() - 1; pos >= 0; pos-- ) {
            targetParag.removeRun(pos);
        }
        //copy新的run
        for(XWPFRun sRun : sourceParag.getRuns()) {
            XWPFRun tarRun = targetParag.createRun();
            copyRun(tarRun, sRun);
        }
    }

	/**
	 * 複製sourceRun至targetRun
	 * @param targetRun
	 * @param sourceRun
	 */
	private void copyRun(XWPFRun targetRun, XWPFRun sourceRun) {
		//設定targetRun屬性
		targetRun.getCTR().setRPr(sourceRun.getCTR().getRPr());
	    targetRun.setText(sourceRun.getText(0));//設定文字
	    List<CTPicture> pictures = sourceRun.getCTR().getPictList();
	    targetRun.getCTR().setPictArray(pictures.toArray(new CTPicture[pictures.size()]));
	}
	
	/**
	 * 從doc找出holder並取代成word,不重複執行
	 * @param doc
	 * @param holder
	 * @param word
	 */
	private void replaceRun(XWPFDocument doc, String holder, String word) {
		replaceRun(doc, holder, word, false);
	}
    
	/**
	 * 從doc找出holder並取代成word,如isMulti為true則重複執行
	 * @param doc
	 * @param holder
	 * @param word
	 */
	private void replaceRun(XWPFDocument doc, String holder, String word,boolean isMulti ) {
		List<XWPFParagraph> p_list = doc.getParagraphs();
		for (XWPFParagraph p : p_list) {
			boolean flag = true ;
			while(flag) {
				flag = replaceRunInPara(p,holder,word ) && isMulti;
			}
		}
	}
    
	/**
	 * 從table找出holder並取代成word,不重複執行
	 * @param table
	 * @param holder
	 * @param word
	 */
	private void replaceRunInTable(XWPFTable table, String holder, String word) {
		replaceRunInTable(table, holder,word, false);
	}
	
	/**
	 * 從table找出holder並取代成word,如isMulti為true 則重複執行
	 * @param table
	 * @param holder
	 * @param word
	 */
	private void replaceRunInTable(XWPFTable table, String holder, String word, boolean isMulti) {
		if (table == null) { return; }
		List<XWPFTableRow> r_list = table.getRows();
		for( int r=0; r <r_list.size()  ;r ++) {
			XWPFTableRow row = r_list.get(r);
			List<XWPFTableCell> c_list = row.getTableCells();
			for( int c = 0; c < c_list.size()  ; c++) {
				XWPFTableCell cell = c_list.get(c);
				for(XWPFParagraph p : cell.getParagraphs()) {
					boolean flag = true ;
					while(flag) {
						flag = replaceRunInPara(p,holder,word ) && isMulti;
					}
				}
			}
		}
	}

	/**
	 * 從 p 找出holder並取代成word,成功取代則回傳true,查無holder則回傳false
	 * @param p
	 * @param holder
	 * @param word
	 * @return
	 */
	private boolean replaceRunInPara(XWPFParagraph p, String holder, String word) {
		if (p.getParagraphText().contains(holder)) {
			int startRun = -1;
	    	int endRun = -1;
	    	List<XWPFRun> r_list = p.getRuns();
	    	for (int pos = 0; pos < r_list.size() ; pos++) {
	    		XWPFRun run = r_list.get(pos);
	    		String run_str = run.getText(0);
		    	if (StringUtils.contains(run_str, holder)) {
	    			run.setText( run_str.replace(holder, word), 0);
	    			return true;
	    		}
	    		if (StringUtils.stripStart(holder,run_str).length() != holder.length()) {
	    			startRun = pos;
					endRun = foundEndInPara(pos, r_list, holder);
					if ( endRun >= 0) { break; }
	    		}else {
					startRun = -1;
					endRun = -1;
	    		}
	    	}
	    	
	    	if (startRun + endRun < 0) { //endRun not found and return 
	    		return false;
	    	}
			
			List<XWPFRun> paragraphRuns = p.getRuns();
			
			String tmp = "";
			for(int i = startRun ; i <= endRun ; i++) {
				tmp += paragraphRuns.get(i).getText(0);
			}
			word = tmp.replace(holder, word);
			
			for (int i = endRun; i > startRun; i--) {
				p.removeRun(i);
			}
			
			XWPFRun paragraphRun = paragraphRuns.get(startRun);
			paragraphRun.setText(word,0);
			return true;
	    }
		return false;
	}
	
	/**
	 * 	從r_list的pos位置為起始點向後找尋該holder最後的Run位置並回傳
	 * @param pos
	 * @param r_list
	 * @param holder
	 * @return
	 */
	private int foundEndInPara(int pos, List<XWPFRun> r_list, String holder) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}");
		String candicate = "";
		for(int j = pos ; j < r_list.size() ; j ++) {
			candicate += r_list.get(j).getText(0);
			Matcher matcher = pattern.matcher(candicate);
			if (matcher.find()) {
				//如果第一個匹配到的${並非holder則跳出
				if (matcher.group().equals(holder) ) {
					return j;
				}
				return -1;
			}
		}
		return -1;
	}
