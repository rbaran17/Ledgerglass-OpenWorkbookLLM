"""Independent AW-061 reference values. Standard-library only; no product imports."""
from fractions import Fraction as F
from itertools import combinations
from statistics import mean, variance, NormalDist
from math import comb, exp, factorial, sqrt
from pathlib import Path
import json, random

def lp(objective, constraints, upper):
    all_constraints = constraints + [[1, 0, upper[0]], [0, 1, upper[1]], [-1, 0, 0], [0, -1, 0]]
    candidates=[]
    for one,two in combinations(all_constraints,2):
        a,b,c=map(F,one); d,e,f=map(F,two); determinant=a*e-b*d
        if not determinant: continue
        x=(c*e-b*f)/determinant; y=(a*f-c*d)/determinant
        if all(F(i)*x+F(j)*y<=F(k) for i,j,k in all_constraints):
            candidates.append((F(objective[0])*x+F(objective[1])*y,x,y))
    if not candidates: return None
    best=max(candidates)
    return {'objective':float(best[0]),'solution':[float(best[1]),float(best[2])], 'exactObjective':str(best[0])}
rng=random.Random(61061)
models=[]
for i in range(24):
    c=[rng.randint(-5,9),rng.randint(-5,9)]
    a=[[rng.randint(1,5),rng.randint(1,5),rng.randint(5,35)] for _ in range(3)]
    upper=[rng.randint(2,10),rng.randint(2,10)]
    models.append({'objective':c,'constraints':a,'upper':upper,'expected':lp(c,a,upper)})
values=[1,2,3,4,5]
result={'generator':'Python standard-library Fraction exhaustive intersections; statistics; NormalDist; exact combinatorial PMF sums. No application code imported.', 'seed':61061,'linearModels':models,'descriptive':{'mean':mean(values),'variance':variance(values),'standardError':sqrt(variance(values)/len(values))},'normal':[{'point':x,'density':NormalDist().pdf(x),'cdf':NormalDist().cdf(x)} for x in [-3,-1,0,1,3]],'binomial':[{'point':x,'mass':comb(10,x)*.3**x*.7**(10-x),'cdf':sum(comb(10,k)*.3**k*.7**(10-k) for k in range(x+1))} for x in [0,1,3,8,10]],'poisson':[{'point':x,'mass':exp(-3)*3**x/factorial(x),'cdf':sum(exp(-3)*3**k/factorial(k) for k in range(x+1))} for x in [0,1,3,8,10]],'norrisCertified':{'source':'https://www.itl.nist.gov/div898/strd/lls/data/LINKS/DATA/Norris.dat','observations':36,'intercept':-.262323073774029,'slope':1.00211681802045,'interceptSE':.232818234301152,'slopeSE':.000429796848199937,'residualSD':.884796396144373,'r2':.999993745883712,'regressionSS':4255954.13232369,'residualSS':26.6173985294224,'df':34,'f':5436385.54079785},'forecast':{'periods':[1,2,3,4],'values':[2,4,6,8],'linear':[10,12],'movingAverageWindow2':[7,7.5],'exponentialAlphaHalf':[6.25,6.25]}}
path=Path(__file__).resolve().parents[1]/'examples/Analysis-independent-reference.json'
path.write_text(json.dumps(result,indent=2)+'\n')
print(json.dumps({'referenceFile':str(path),'independentLinearModels':len(models)}))

# High-precision independent beta integral by its power series for B(17, 1/2).
# (1-t)^(-1/2) = sum(comb(2*k,k)/4**k * t**k); integrate term by term.
from decimal import Decimal, localcontext
with localcontext() as context:
    context.prec=80
    df=Decimal(34)
    x=df/(df+Decimal(str(result['norrisCertified']['f'])))
    beta=Decimal(4**17*factorial(17)*factorial(16))/Decimal(factorial(34))
    term=x**17
    total=term/Decimal(17)
    coefficient=Decimal(1)
    for k in range(1,1000):
        coefficient*=Decimal(2*k-1)/Decimal(2*k)
        term*=x
        addition=coefficient*term/Decimal(17+k)
        total+=addition
        if abs(addition)<abs(total)*Decimal('1e-70'):break
    result['norrisCertified']['tailProbabilityDecimal']=str(total/beta)
    result['norrisCertified']['tailProbabilityMethod']='80-digit Decimal beta-integral power series with exact rational B(17,1/2), based on certified F.'
path.write_text(json.dumps(result,indent=2)+'\n')
with localcontext() as context:
    context.prec=80
    t=Decimal(str(result['norrisCertified']['intercept']))/Decimal(str(result['norrisCertified']['interceptSE']))
    x=Decimal(34)/(Decimal(34)+t*t)
    beta=Decimal(4**17*factorial(17)*factorial(16))/Decimal(factorial(34))
    power=x**17; coefficient=Decimal(1); total=power/Decimal(17)
    for k in range(1,10000):
        coefficient*=Decimal(2*k-1)/Decimal(2*k);power*=x
        addition=coefficient*power/Decimal(17+k);total+=addition
        if abs(addition)<abs(total)*Decimal('1e-70'):break
    result['norrisCertified']['interceptPDecimal']=str(total/beta)
path.write_text(json.dumps(result,indent=2)+'\n')
