# -*- coding: utf-8 -*-
"""
gamma_pure.py
=================================================================
목적
----
표본크기 산정(PoissonSample) 및 통계적 표본결과평가(신뢰계수/증분계수,
GAMMA.INV 방식)에 필요한 "정칙화 불완전감마함수(Regularized Incomplete
Gamma Function)"를 SciPy 없이 순수 파이썬(표준 라이브러리 math만 사용)
으로 구현한다.

왜 SciPy를 안 쓰는가
--------------------
PyInstaller로 단일 실행파일(.exe)을 만들 때 SciPy를 포함시키면 실행파일
용량이 수십 MB~100MB 이상으로 크게 늘어난다. 반면 아래 구현은 표준
라이브러리(math)만 사용하므로 실행파일이 훨씬 가볍다.

구현 알고리즘 (Numerical Recipes 표준 기법)
-------------------------------------------
1. gammap(a, x): 정칙화 하부불완전감마함수 P(a,x)
   - x < a+1 이면 급수전개(Series expansion) 사용
   - x >= a+1 이면 연분수(Continued Fraction) 사용
   (Press, Teukolsky, Vetterling, Flannery, "Numerical Recipes", 6.2절)

2. gammap_inv(a, p): P(a,x)=p 를 만족하는 x를 구하는 역함수
   - 초기값 추정 후 뉴턴법(Newton's method)으로 수렴
   - gammap()이 이미 검증되었으므로, 그 함수의 근을 찾는 방식이라 정확도는
     gammap() 자체의 정확도(더블 정밀도, 소수점 15자리 내외)에 좌우됨

검증
----
scipy.special.gammainc / gammaincinv 결과와 대조 검증하였으며, 그 결과는
별도 보고서(통계적/비통계적 각 1부)로 정리되어 있다.
"""

import math


def _log_gamma(x):
    """ln(Gamma(x)) - Lanczos 근사 (g=7, n=9), 배정밀도 기준 오차 매우 작음."""
    g = 7
    coeffs = [
        0.99999999999980993, 676.5203681218851, -1259.1392167224028,
        771.32342877765313, -176.61502916214059, 12.507343278686905,
        -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
    ]
    if x < 0.5:
        # 반사공식(Reflection formula): Gamma(x)*Gamma(1-x) = pi/sin(pi*x)
        return math.log(math.pi / math.sin(math.pi * x)) - _log_gamma(1 - x)
    x -= 1
    a = coeffs[0]
    t = x + g + 0.5
    for i in range(1, g + 2):
        a += coeffs[i] / (x + i)
    return 0.5 * math.log(2 * math.pi) + (x + 0.5) * math.log(t) - t + math.log(a)


def _gammap_series(a, x):
    """P(a,x) 급수전개 (x < a+1 구간에서 사용)."""
    if x <= 0:
        return 0.0
    ap = a
    total = 1.0 / a
    delta = total
    for _ in range(500):
        ap += 1
        delta *= x / ap
        total += delta
        if abs(delta) < abs(total) * 1e-15:
            break
    return total * math.exp(-x + a * math.log(x) - _log_gamma(a))


def _gammap_cf(a, x):
    """Q(a,x)=1-P(a,x) 연분수 전개 (x >= a+1 구간에서 사용, 반환은 P(a,x))."""
    tiny = 1e-300
    b = x + 1 - a
    c = 1 / tiny
    d = 1 / b
    h = d
    for i in range(1, 500):
        an = -i * (i - a)
        b += 2
        d = an * d + b
        if abs(d) < tiny:
            d = tiny
        c = b + an / c
        if abs(c) < tiny:
            c = tiny
        d = 1 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < 1e-15:
            break
    q = math.exp(-x + a * math.log(x) - _log_gamma(a)) * h
    return 1.0 - q


def gammap(a, x):
    """정칙화 하부불완전감마함수 P(a,x) = γ(a,x)/Γ(a).

    Excel의 GAMMADIST(x, a, 1, TRUE) 와 동일한 값을 반환한다.
    """
    if x < 0 or a <= 0:
        raise ValueError("gammap: a>0, x>=0 이어야 합니다.")
    if x == 0:
        return 0.0
    if x < a + 1:
        return _gammap_series(a, x)
    else:
        return _gammap_cf(a, x)


def gammap_inv(a, p, x0=None):
    """P(a,x) = p 를 만족하는 x 를 뉴턴법으로 구한다.

    Excel의 GAMMA.INV(p, a, 1) 과 동일한 값을 반환한다.
    """
    if not (0 < p < 1):
        raise ValueError("gammap_inv: 0<p<1 이어야 합니다.")

    # 초기값: Wilson-Hilferty 근사 (카이제곱분포 근사 기반)
    if x0 is None:
        if a > 1:
            pp = p if p < 0.5 else 1 - p
            t = math.sqrt(-2 * math.log(pp))
            x = t - (2.30753 + 0.27061 * t) / (1 + (0.99229 + 0.04481 * t) * t)
            if p < 0.5:
                x = -x
            x = max(
                1e-3,
                a * (1 - 1.0 / (9 * a) + x * math.sqrt(1.0 / (9 * a))) ** 3,
            )
        else:
            t = 1.0 - a * (0.253 + a * 0.12)
            if p < t:
                x = (p / t) ** (1.0 / a)
            else:
                x = 1 - math.log(1 - (p - t) / (1 - t))
    else:
        x = x0

    log_gamma_a = _log_gamma(a)
    for _ in range(100):
        if x <= 0:
            x = 1e-6
        err = gammap(a, x) - p
        # dP/dx = x^(a-1) * e^-x / Gamma(a)
        deriv = math.exp((a - 1) * math.log(x) - x - log_gamma_a)
        if deriv == 0:
            break
        delta = err / deriv
        x_new = x - delta
        if x_new <= 0:
            x_new = x / 2
        if abs(x_new - x) < 1e-12 * max(1.0, x):
            x = x_new
            break
        x = x_new
    return x


# ----------------------------------------------------------------------
# 감사 표본추출 응용 함수
# ----------------------------------------------------------------------
def poisson_sample(risk, pE, pT):
    """포아송분포 기반 표본크기 산정 (매크로 PoissonSample 함수와 동일 로직).

    risk : 발견위험 또는 신뢰상실위험 (예: 0.05)
    pE   : 예상왜곡표시율 (모집단 대비 비율)
    pT   : 허용오류율 (모집단 대비 비율)
    """
    if risk <= 0 or risk >= 1 or pE < 0 or pE >= 1 or pT <= 0 or pT >= 1:
        raise ValueError("poisson_sample: 입력값 범위를 벗어났습니다.")
    n = math.ceil(-math.log(risk) / pT)
    while n <= 5_000_000:
        if gammap(1 + pE * n, n * pT) >= 1 - risk:
            break
        n += 1
    return n


def reliability_factor(confidence, n_errors):
    """n_errors개의 오류를 가정했을 때의 누적 포아송 신뢰계수.

    Excel: GAMMA.INV(confidence, n_errors+1, 1) 과 동일
    """
    return gammap_inv(n_errors + 1, confidence)


def incremental_factor(confidence, n_errors):
    """n_errors번째(1부터 시작) 오류의 증분계수 = R(n) - R(n-1)."""
    return reliability_factor(confidence, n_errors) - reliability_factor(confidence, n_errors - 1)
