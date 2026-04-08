from __future__ import annotations

import shutil
import sys
from pathlib import Path

from docx import Document

H1 = "Codex Heading 1"
H2 = "Codex Heading 2"


SECTIONS = [
    ("pagebreak", ""),
    ("heading1", "学位论文答辩委员会组成和答辩决议"),
    ("p", "答辩委员会组成："),
    ("p", "主席："),
    ("p", "委员："),
    ("p", "秘书："),
    ("p", "答辩时间："),
    ("p", "答辩地点："),
    ("p", "答辩决议："),
    ("p", "此处请根据学校最终答辩委员会决议表格内容回填。"),
    ("pagebreak", ""),
    ("heading1", "攻读学位期间获得学术成果情况"),
    ("heading2", "攻读博士学位期间发表的主要学术论文"),
    ("p", "1. Deng G, Fu L. A bedaquiline, pretomanid, moxifloxacin, and pyrazinamide regimen for drug-susceptible and drug-resistant tuberculosis[J]. Lancet Infect Dis, 2024, 24(9): 940-941."),
    ("p", "2. Fu L, Wang W, Xiong J, et al. Evaluation of Sulfasalazine as an Adjunctive Therapy in Treating Pulmonary Pre-XDR-TB: Efficacy, Safety, and Treatment Implication[J]. Infect Drug Resist, 2024, 17: 595-604."),
    ("p", "3. Fu L, Xiong J, Wang H, et al. Study protocol for safety and efficacy of all-oral shortened regimens for multidrug-resistant tuberculosis: a multicenter randomized withdrawal trial and a single-arm trial [SEAL-MDR][J]. BMC Infect Dis, 2023, 23(1): 834."),
    ("p", "4. Fu L, Zhang X, Xiong J, et al. Selecting an appropriate all-oral short-course regimen for patients with multidrug-resistant or pre-extensive drug-resistant tuberculosis in China: a multicenter prospective cohort study[J]. Int J Infect Dis, 2023, 135: 101-108."),
    ("p", "5. Fu L, Weng T, Sun F, et al. Insignificant difference in culture conversion between bedaquiline-containing and bedaquiline-free all-oral short regimens for multidrug-resistant tuberculosis[J]. Int J Infect Dis, 2021, 111: 138-147."),
    ("p", "6. Fu L, Feng Y, Ren T, et al. Detecting latent tuberculosis infection with a breath test using mass spectrometer: a pilot cross-sectional study[J]. Biosci Trends, 2023, 17(1): 73-77."),
    ("p", "7. Fu L, Wang L, Wang H, et al. A cross-sectional study: a breathomics based pulmonary tuberculosis detection method[J]. BMC Infect Dis, 2023, 23(1): 148."),
    ("p", "8. 付亮，邓国防. MDR-Chin研究解析：耐多药肺结核全口服短程治疗方案在中国的应用前景[J]. 中国防痨杂志, 2024, 46(1): 18-22。"),
    ("heading2", "攻读博士学位期间承担或参与的主要科研课题"),
    ("p", "1. 2025年度新发突发与重大传染病防控国家科技重大专项计划“结核病诊断突破性新技术研究（技术开发）”项目“基于宿主转录标志物的结核病筛查技术的研究”子任务，子任务负责人。"),
    ("p", "2. 2025年度新发突发与重大传染病防控国家科技重大专项计划“糖尿病人群结核病短程治疗方案研究”，参与人。"),
    ("p", "3. 2024年深圳市医学研究专项资金临床多中心研究“泛用性超短程方案治疗药物敏感性和耐药肺结核的多中心随机对照试验”，参与人。"),
    ("p", "4. 2020年度国家自然科学基金面上项目“宿主SNP对柳氮磺吡啶调控巨噬细胞结核免疫效果的影响及其机制研究”，第二负责人。"),
    ("p", "5. 中国耐多药结核病超短程方案研究：全口服短程方案治疗氟喹诺酮类敏感性耐多药结核病的多中心随机撤药试验，第二负责人。"),
    ("p", "6. 基于新型呼出气高通量实时质谱检测的感染性肺炎诊断和疗效评估的前瞻性队列研究，负责人。"),
    ("p", "7. 佛山市第四人民医院登峰计划开放课题“中国耐多药肺结核全口服短程抗结核方案多中心临床研究”，第二负责人。"),
    ("p", "8. 深圳市科创委自由探索项目“复治耐多药结核病患者MTB耐药谱变迁的分子机制研究”，负责人。"),
    ("heading2", "其他学术与专业发展成果"),
    ("p", "1. 入选国家卫生健康委员会司“全球卫生人才后备库”。"),
    ("p", "2. 入选中国结核病临床试验合作中心首批（2016年）“青年医师临床科研培训”项目。"),
    ("p", "3. 获得NEJM首届（2024年）“高水平临床研究培训认证项目”认证。"),
    ("p", "4. 入选北京结核病诊疗技术创新联盟“雏鹰计划”及深圳市第三人民医院“优才计划”。"),
    ("pagebreak", ""),
    ("heading1", "个人简历"),
    ("p", "付亮，男，汉族，1981年7月生，湖南人，中共党员，主治医师。"),
    ("p", "2000年9月至2005年6月就读于南方医科大学临床医学专业，获学士学位；2005年9月至2008年6月就读于南方医科大学呼吸病学专业，获硕士学位，师从蔡绍曦教授；2017年9月至今在首都医科大学攻读博士学位，师从唐神结教授。"),
    ("p", "2008年7月至2012年5月在广东医学院附属深圳市第三人民医院肺二科任住院医师；2012年5月至今在广东医学院附属深圳市第三人民医院肺二科任主治医师；2017年9月至2018年3月在北京胸科医院进修。"),
    ("p", "长期从事结核病、耐药结核病及非结核分枝杆菌病的临床诊疗与研究工作，研究方向主要包括结核病快速无创诊断、耐药结核病短程治疗以及非结核分枝杆菌病治疗新方案。"),
]


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: python generate_final_thesis_docx_appended.py <source.docx> <target.docx>")
        return 1

    src = Path(sys.argv[1])
    target = Path(sys.argv[2])
    if not src.exists():
        print(f"Source not found: {src}")
        return 1

    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, target)
    doc = Document(str(target))

    for kind, text in SECTIONS:
        if kind == "pagebreak":
            doc.add_page_break()
        else:
            p = doc.add_paragraph(text)
            if kind == "heading1":
                p.style = H1
            elif kind == "heading2":
                p.style = H2

    doc.save(str(target))
    print(f"Generated appended final draft: {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
