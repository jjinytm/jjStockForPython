"""
공통함수 모음
"""
__all__ = ["_get_connection_state"]


# 통신접속상태 확인
def _get_connection_state(kiwoom):
    state = kiwoom.dynamicCall("GetConnectState()")
    return state
