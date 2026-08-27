/**
 * 지킴이 캘린더 → 근무일지 시트 연동 (cal.html의 [저장 + 시트 전송])
 *
 * 설치 (근무일지 스프레드시트에서, 2분):
 *  1) 시트 열기 → 확장 프로그램 → Apps Script → 이 파일 내용 전체 붙여넣기 → 저장
 *  2) 배포 → 새 배포 → 유형: 웹 앱 → 실행 계정: 나 / 액세스 권한: 모든 사용자 → 배포
 *  3) 웹 앱 URL 복사 → cal.html의 GAS_URL 상수에 붙여넣기 (클코에 "GAS URL 이거야" 한 줄이면 커밋해줌)
 *
 * 시트 구조 전제: 'wk N' 행의 B~H열에 일(day) 숫자, 그 아래 블록에
 * 일방문자수/근무자/단체관람예약/참가자수/특이사항 라벨 행 (행 순서 무관, 라벨로 찾음)
 */
var TAB = '전시기간 주요 행사 일정표';

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var parts = String(d.date || '').split('-'); // 2026-09-05
    if (parts.length !== 3 || parts[1] !== '09') return out_({ok:false, err:'9월 날짜만 시트에 있음: ' + d.date});
    var day = Number(parts[2]);

    var sh = SpreadsheetApp.getActive().getSheetByName(TAB);
    if (!sh) return out_({ok:false, err:'탭 없음: ' + TAB});
    var vals = sh.getDataRange().getValues();

    // 해당 날짜의 wk 블록·요일 열 찾기
    var wkRow = -1, col = -1;
    for (var r = 0; r < vals.length; r++) {
      if (String(vals[r][0]).trim().toLowerCase().indexOf('wk') === 0) {
        for (var c = 1; c <= 7; c++) {
          if (Number(vals[r][c]) === day) { wkRow = r; col = c; break; }
        }
        if (wkRow >= 0) break;
      }
    }
    if (wkRow < 0) return out_({ok:false, err:'시트에서 ' + day + '일을 못 찾음'});

    // 블록 내 라벨 행 매핑 (다음 wk 전까지)
    var rowOf = {};
    for (var r2 = wkRow + 1; r2 < vals.length; r2++) {
      var label = String(vals[r2][0]).trim();
      if (label.toLowerCase().indexOf('wk') === 0) break;
      if (label) rowOf[label] = r2;
    }
    function setCell(label, text) {
      if (text == null || rowOf[label] == null) return;
      sh.getRange(rowOf[label] + 1, col + 1).setValue(text);
    }

    if (d.visitors)     setCell('일방문자수', '실제: ' + (d.visitors.actual || '') + '명\n보고용: ' + (d.visitors.reported || '') + '명');
    if (d.participants) setCell('참가자수', '사전신청: ' + (d.participants.pre || '') + '\n현장참여: ' + (d.participants.onsite || ''));
    if (d.workers != null) setCell('근무자', d.workers);
    if (d.group != null)   setCell('단체관람예약', d.group);
    if (d.notes != null)   setCell('특이사항', d.notes);

    return out_({ok:true});
  } catch (err) {
    return out_({ok:false, err:String(err)});
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
