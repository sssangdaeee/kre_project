
"""
death_rate_ml_forecast.py

목표
- 업로드된 데이터(2010-2024)를 이용해 death_rate를 예측하는 3개 모델(DecisionTree, RandomForest, XGBoost) 학습
- 2010-2018 학습 / 2019-2024 테스트로 성능 비교 (가중치= count 노출치)
- 각 모델별 중요 변수 계산 및 저장
- 가장 성능이 좋은 모델을 2010-2024 전체로 재학습 후, 2025-2035년 death_rate 예측
  (가정: 2025-2035의 설명변수는 2024년 값이 유지됨; year_std도 2024값으로 고정)

입출력
- 입력: ML_Input_v1.xlsx (컬럼: year_std, age_group, F_ratio, M_ratio, nonblood_ratio, blood_ratio, NS1_ratio, NS2_ratio, NS3_ratio, NS4_ratio, SM1_ratio, SM2_ratio, unknown_ratio, count, death_count, death_rate)
- 출력: ./outputs/ 디렉토리에 다음 파일 생성
  - metrics_by_model.csv            : 모델별 성능 지표 (가중 RMSE/MAE/MAPE, R2)
  - feature_importance_<model>.csv  : 각 모델의 중요변수
  - best_model_name.txt             : 선택된 베스트 모델명
  - forecast_2025_2035.csv          : 2025-2035 연도별/연령대별 예측 death_rate
  - model_best.joblib               : 학습된 베스트 모델(pipeline 포함)
  - top_features_<model>.png        : 상위 20개 중요변수 바차트 (모델별)

사용
- python death_rate_ml_forecast.py --data /path/to/ML_Input_v1.xlsx
- 필요한 패키지: pandas, numpy, scikit-learn, xgboost, matplotlib, joblib

주의
- 일부 연령/연도에서 death_count가 작아 death_rate 변동성이 큼 ⇒ count(노출치)로 가중 학습
- age_group은 범주형으로 원-핫 인코딩
- 학습/평가/재학습/예측 전 과정에서 같은 전처리 파이프라인 사용
"""

import argparse
import os
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd

from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder
from sklearn.impute import SimpleImputer
from sklearn.pipeline import Pipeline
from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score
from sklearn.utils.validation import check_is_fitted

import matplotlib.pyplot as plt

try:
    from xgboost import XGBRegressor
    _XGB_AVAILABLE = True
except Exception as e:
    print("[경고] xgboost 임포트 실패:", e)
    print("       XGBoost 모델은 건너뜁니다. (pip install xgboost 로 설치 가능)")
    _XGB_AVAILABLE = False

from joblib import dump

RANDOM_STATE = 42

def weighted_mape(y_true, y_pred, sample_weight=None, eps=1e-9):
    # MAPE with stabilization for zeros, and optional weights
    y_true = np.asarray(y_true)
    y_pred = np.asarray(y_pred)
    denom = np.maximum(np.abs(y_true), eps)
    mape = np.abs((y_true - y_pred) / denom)
    if sample_weight is None:
        return float(np.mean(mape) * 100.0)
    sw = np.asarray(sample_weight)
    return float(np.average(mape, weights=sw) * 100.0)

def weighted_rmse(y_true, y_pred, sample_weight=None):
    if sample_weight is None:
        return float(np.sqrt(mean_squared_error(y_true, y_pred)))
    return float(np.sqrt(np.average((y_true - y_pred) ** 2, weights=sample_weight)))

def build_preprocessor(cat_cols: List[str], num_cols: List[str]) -> ColumnTransformer:
    """ColumnTransformer: 범주형(최빈값 대치 + OHE), 수치형(중앙값 대치 + 패스스루)"""
    categorical_pipeline = Pipeline(steps=[
        ("imputer", SimpleImputer(strategy="most_frequent")),
        ("ohe", OneHotEncoder(handle_unknown="ignore", sparse=False))
    ])
    numeric_pipeline = Pipeline(steps=[
        ("imputer", SimpleImputer(strategy="median"))
    ])
    preprocessor = ColumnTransformer(
        transformers=[
            ("cat", categorical_pipeline, cat_cols),
            ("num", numeric_pipeline, num_cols)
        ],
        remainder="drop"
    )
    return preprocessor

def get_feature_names(preprocessor: ColumnTransformer, cat_cols: List[str], num_cols: List[str]) -> List[str]:
    """전처리 후 최종 피처명 반환 (sklearn >=1.0 권장)"""
    feature_names = []
    try:
        # sklearn 1.0+: get_feature_names_out 지원
        ohe = preprocessor.named_transformers_["cat"].named_steps["ohe"]
        cat_out = ohe.get_feature_names_out(cat_cols).tolist()
        num_out = num_cols  # numeric pipeline doesn't change names
        feature_names = cat_out + num_out
    except Exception:
        # Fallback: 대략적 네이밍
        ohe = preprocessor.named_transformers_["cat"].named_steps["ohe"]
        cat_out = []
        for i, c in enumerate(cat_cols):
            categories = getattr(ohe, "categories_", [])[i] if hasattr(ohe, "categories_") else []
            if len(categories) == 0:
                cat_out.append(f"{c}_encoded")
            else:
                cat_out.extend([f"{c}_{v}" for v in categories])
        num_out = num_cols
        feature_names = cat_out + num_out
    return feature_names

def train_and_eval_model(model_name: str,
                         estimator,
                         X_train: pd.DataFrame, y_train: pd.Series, w_train: np.ndarray,
                         X_test: pd.DataFrame, y_test: pd.Series, w_test: np.ndarray,
                         cat_cols: List[str], num_cols: List[str]) -> Tuple[Pipeline, Dict[str, float], pd.DataFrame]:
    """단일 모델 학습/평가 및 중요변수 산출"""
    preprocessor = build_preprocessor(cat_cols, num_cols)
    pipe = Pipeline(steps=[("prep", preprocessor),
                          ("model", estimator)])

    pipe.fit(X_train, y_train, model__sample_weight=w_train)

    y_pred = pipe.predict(X_test)
    metrics = {
        "model": model_name,
        "weighted_RMSE": weighted_rmse(y_test, y_pred, w_test),
        "weighted_MAE": float(np.average(np.abs(y_test - y_pred), weights=w_test)),
        "weighted_MAPE(%)": weighted_mape(y_test, y_pred, w_test),
        "R2": float(r2_score(y_test, y_pred, sample_weight=w_test))
    }

    # Feature importance 추출
    feature_importance_df = pd.DataFrame()
    try:
        # 파이프라인에서 전처리기와 모델에 접근
        prep = pipe.named_steps["prep"]
        model = pipe.named_steps["model"]
        check_is_fitted(prep)
        check_is_fitted(model)

        feature_names = get_feature_names(prep, cat_cols, num_cols)

        if hasattr(model, "feature_importances_"):
            importances = model.feature_importances_
            feature_importance_df = pd.DataFrame({
                "feature": feature_names[:len(importances)],
                "importance": importances
            }).sort_values("importance", ascending=False)
        elif hasattr(model, "coef_"):
            coefs = np.ravel(model.coef_)
            feature_importance_df = pd.DataFrame({
                "feature": feature_names[:len(coefs)],
                "importance": np.abs(coefs)
            }).sort_values("importance", ascending=False)
        else:
            feature_importance_df = pd.DataFrame(columns=["feature", "importance"])
    except Exception as e:
        print(f"[경고] 중요변수 계산 실패 ({model_name}): {e}")
        feature_importance_df = pd.DataFrame(columns=["feature", "importance"])

    return pipe, metrics, feature_importance_df

def freeze_covariates_from_2024(df: pd.DataFrame,
                                feature_cols: List[str],
                                freeze_year_col: str = "year_std",
                                ref_year: int = 2024,
                                forecast_years: List[int] = list(range(2025, 2036)),
                                id_cols: List[str] = ["age_group"]) -> pd.DataFrame:
    """2024년의 공변량을 그대로 복제하여 미래 연도(2025-2035)용 피처셋 생성"""
    if freeze_year_col not in df.columns:
        raise ValueError(f"'{freeze_year_col}' 컬럼이 데이터에 없습니다.")

    # 2024 기준 템플릿
    base_2024 = df[df[freeze_year_col] == ref_year].copy()
    if base_2024.empty:
        raise ValueError("2024년 데이터가 없습니다. 2024년 값으로 동결(freeze)할 수 없습니다.")

    # 필요한 컬럼만 유지(피처 + ID)
    keep_cols = list(dict.fromkeys(id_cols + feature_cols))  # preserve order & unique
    base_2024 = base_2024[keep_cols].copy()

    # 미래 연도별로 복제
    out_list = []
    for yr in forecast_years:
        tmp = base_2024.copy()
        # year_std(또는 time proxy)는 2024값을 유지(동결)
        # 단, 결과 표에는 예측대상 연도를 별도 컬럼 'forecast_year'로 기록
        tmp["forecast_year"] = yr
        out_list.append(tmp)

    future_df = pd.concat(out_list, axis=0, ignore_index=True)
    return future_df

def plot_top_importances(imp_df: pd.DataFrame, model_name: str, out_dir: str, top_k: int = 20):
    if imp_df is None or imp_df.empty:
        return
    top = imp_df.head(top_k)
    plt.figure()
    plt.barh(top["feature"][::-1], top["importance"][::-1])
    plt.title(f"Top {top_k} Feature Importances - {model_name}")
    plt.xlabel("Importance")
    plt.ylabel("Feature")
    plt.tight_layout()
    path = os.path.join(out_dir, f"top_features_{model_name}.png")
    plt.savefig(path, dpi=150)
    plt.close()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", type=str, default="ML_Input_v1.xlsx", help="입력 데이터 경로 (엑셀)")
    parser.add_argument("--outdir", type=str, default="outputs", help="결과 저장 디렉토리")
    args = parser.parse_args()

    os.makedirs(args.outdir, exist_ok=True)

    # -----------------------------
    # 1) 데이터 로드
    # -----------------------------
    df = pd.read_excel(args.data)

    expected_cols = [
        "year_std", "age_group",
        "F_ratio", "M_ratio", "nonblood_ratio", "blood_ratio",
        "NS1_ratio", "NS2_ratio", "NS3_ratio", "NS4_ratio",
        "SM1_ratio", "SM2_ratio", "unknown_ratio",
        "count", "death_count", "death_rate"
    ]
    missing = [c for c in expected_cols if c not in df.columns]
    if missing:
        print("[경고] 예상 컬럼 중 누락:", missing)
        # 계속 진행하되, 가능한 컬럼만 사용

    # -----------------------------
    # 2) 피처/타겟/가중치 정의
    # -----------------------------
    target_col = "death_rate"
    weight_col = "count" if "count" in df.columns else None

    # 예측에 사용하는 피처(의도적으로 'count', 'death_count', 'death_rate' 제외)
    base_feature_cols = [
        c for c in [
            "year_std", "age_group",
            "F_ratio", "M_ratio", "nonblood_ratio", "blood_ratio",
            "NS1_ratio", "NS2_ratio", "NS3_ratio", "NS4_ratio",
            "SM1_ratio", "SM2_ratio", "unknown_ratio"
        ] if c in df.columns
    ]

    # 범주형/수치형 분할
    cat_cols = [c for c in ["age_group"] if c in base_feature_cols]
    num_cols = [c for c in base_feature_cols if c not in cat_cols]

    # 학습/평가 기간 필터
    if "year_std" not in df.columns:
        raise ValueError("year_std 컬럼이 필요합니다.")
    train_mask = (df["year_std"] >= 2010) & (df["year_std"] <= 2018)
    test_mask  = (df["year_std"] >= 2019) & (df["year_std"] <= 2024)

    df_train = df.loc[train_mask].copy()
    df_test  = df.loc[test_mask].copy()

    # 타겟/가중치
    y_train = df_train[target_col].astype(float)
    y_test  = df_test[target_col].astype(float)
    w_train = df_train[weight_col].clip(lower=1).astype(float).to_numpy() if weight_col else None
    w_test  = df_test[weight_col].clip(lower=1).astype(float).to_numpy() if weight_col else None

    X_train = df_train[base_feature_cols].copy()
    X_test  = df_test[base_feature_cols].copy()

    # -----------------------------
    # 3) 모델 정의
    # -----------------------------
    models = {
        "DecisionTree": DecisionTreeRegressor(
            random_state=RANDOM_STATE,
            min_samples_leaf=5,  # 작은 수의 관측치에 대한 과적합 방지
            max_depth=None
        ),
        "RandomForest": RandomForestRegressor(
            n_estimators=400,
            random_state=RANDOM_STATE,
            n_jobs=-1,
            min_samples_leaf=3
        ),
    }
    if _XGB_AVAILABLE:
        models["XGBoost"] = XGBRegressor(
            n_estimators=800,
            max_depth=4,
            learning_rate=0.05,
            subsample=0.8,
            colsample_bytree=0.8,
            reg_lambda=1.0,
            random_state=RANDOM_STATE,
            n_jobs=-1,
            tree_method="hist"
        )

    # -----------------------------
    # 4) 학습/성능평가/중요변수
    # -----------------------------
    all_metrics = []
    all_importances = {}

    for name, est in models.items():
        pipe, metrics, imp_df = train_and_eval_model(
            model_name=name,
            estimator=est,
            X_train=X_train, y_train=y_train, w_train=w_train,
            X_test=X_test, y_test=y_test, w_test=w_test,
            cat_cols=cat_cols, num_cols=num_cols
        )
        all_metrics.append(metrics)
        all_importances[name] = imp_df

        # 저장
        imp_path = os.path.join(args.outdir, f"feature_importance_{name}.csv")
        imp_df.to_csv(imp_path, index=False)

        # 중요변수 그림
        plot_top_importances(imp_df, name, args.outdir, top_k=20)

        # 파이프라인 자체도 임시 저장(원하면 비교 분석에 사용)
        dump(pipe, os.path.join(args.outdir, f"model_{name}.joblib"))

    metrics_df = pd.DataFrame(all_metrics).sort_values("weighted_RMSE")
    metrics_path = os.path.join(args.outdir, "metrics_by_model.csv")
    metrics_df.to_csv(metrics_path, index=False)

    print("\n=== 모델 성능(2010-2018 학습 / 2019-2024 테스트) ===")
    print(metrics_df.to_string(index=False))

    # -----------------------------
    # 5) 베스트 모델 선택(가중 RMSE 최소)
    # -----------------------------
    best_row = metrics_df.iloc[0]
    best_name = str(best_row["model"])
    with open(os.path.join(args.outdir, "best_model_name.txt"), "w") as f:
        f.write(best_name + "\n")
    print(f"\n[베스트 모델] {best_name}")

    # -----------------------------
    # 6) 2010-2024 전체로 재학습
    # -----------------------------
    train_all_mask = (df["year_std"] >= 2010) & (df["year_std"] <= 2024)
    df_all = df.loc[train_all_mask].copy()
    X_all = df_all[base_feature_cols].copy()
    y_all = df_all[target_col].astype(float)
    w_all = df_all[weight_col].clip(lower=1).astype(float).to_numpy() if weight_col else None

    # 베스트 모델 인스턴스 재생성(동일 하이퍼파라미터)
    if best_name == "DecisionTree":
        best_estimator = DecisionTreeRegressor(random_state=RANDOM_STATE, min_samples_leaf=5, max_depth=None)
    elif best_name == "RandomForest":
        best_estimator = RandomForestRegressor(n_estimators=400, random_state=RANDOM_STATE, n_jobs=-1, min_samples_leaf=3)
    else:
        if not _XGB_AVAILABLE:
            raise RuntimeError("베스트 모델이 XGBoost로 선정되었으나 xgboost가 설치되어 있지 않습니다.")
        best_estimator = XGBRegressor(
            n_estimators=800, max_depth=4, learning_rate=0.05,
            subsample=0.8, colsample_bytree=0.8, reg_lambda=1.0,
            random_state=RANDOM_STATE, n_jobs=-1, tree_method="hist"
        )

    best_preprocessor = build_preprocessor(cat_cols, num_cols)
    best_pipe = Pipeline(steps=[("prep", best_preprocessor),
                               ("model", best_estimator)])
    best_pipe.fit(X_all, y_all, model__sample_weight=w_all)

    # 저장
    dump(best_pipe, os.path.join(args.outdir, "model_best.joblib"))

    # -----------------------------
    # 7) 2025-2035 예측 (2024년 공변량 동결)
    # -----------------------------
    feature_cols_for_freeze = base_feature_cols.copy()  # 모델에 들어가는 피처들
    # 미래 피처셋 생성 (year_std 포함 피처는 2024값 유지됨)
    future_X = freeze_covariates_from_2024(
        df=df_all,  # 2010-2024 범위(여기엔 2024가 포함)
        feature_cols=feature_cols_for_freeze,
        freeze_year_col="year_std",
        ref_year=2024,
        forecast_years=list(range(2025, 2036)),
        id_cols=["age_group"]  # 기준 ID 축
    )

    # best_pipe는 'forecast_year'를 사용하지 않으므로 입력에서 제외
    preds = best_pipe.predict(future_X[base_feature_cols])
    future_out = future_X[["age_group", "forecast_year"]].copy()
    future_out["pred_death_rate"] = preds

    # 결과 저장
    forecast_path = os.path.join(args.outdir, "forecast_2025_2035.csv")
    future_out.to_csv(forecast_path, index=False)

    print(f"\n[완료] 예측 결과 저장: {forecast_path}")
    print(f"[참고] 베스트 모델: {best_name}")
    print(f"[참고] 성능표 저장: {metrics_path}")
    print(f"[참고] 중요변수 저장: {args.outdir}/feature_importance_<model>.csv")
    print(f"[참고] 중요변수 TOP20 그림: {args.outdir}/top_features_<model>.png")

if __name__ == "__main__":
    main()
