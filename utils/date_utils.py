from datetime import datetime, date, timedelta
import holidays

def get_last_workday_str():
    """
    한국의 마지막 근무일을 반환하는 함수
    
    Returns:
        str: 마지막 근무일 문자열 (YY.MM.DD. 요일)
    """
    kr_holidays = holidays.KR(years=date.today().year)
    day = date.today() - timedelta(days=1)
    while day in kr_holidays or day.weekday() >= 5:
        day -= timedelta(days=1)

    weekdays_kr = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']
    day_str = day.strftime("%y.%m.%d.")
    weekday_str = weekdays_kr[day.weekday()]
    return f"{day_str} {weekday_str}"

def get_output_filename(today_str, index=None):
    """
    출력 파일 이름 생성 함수
    
    Args:
        today_str: 오늘 날짜 문자열
        index: 파일 인덱스 (None이면 기본 파일명)
    
    Returns:
        str: 출력 파일명
    """
    base = f"({today_str}_대상) 현장 보수교육 대체 사이버 교육 이수자 리스트"
    return f"{base}.xlsx" if index is None else f"{base}_{index}.xlsx"