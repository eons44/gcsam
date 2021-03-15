from plantManagement import *
import matplotlib.pyplot as plt
import os
import math

class GraphGenerator:
    m_lineManager = None
    m_directory = "graphs"
    m_cachedPlants = []

    def __init__(self, lineManager, directory="graphs"):
        self.m_lineManager = lineManager
        self.m_directory = directory
        
        #Assume ctor is called from working directory or absolute path.
        if not os.path.exists(directory):
            os.makedirs(directory)

    def GetPlants(self):
        if (not len(self.m_cachedPlants)):
            self.m_cachedPlants = [p for l in self.m_lineManager.m_lines for p in l.m_plants]
        return self.m_cachedPlants

    def GenerateGraphForTotalWeightFraction(self):
        # plt.style.use('ggplot')
        plants = [p.m_name for p in self.GetPlants()]
        twf = [p.m_totalWeightFraction for p in self.GetPlants()]

        x_pos = [i for i, _ in enumerate(x)]

        plt.bar(x_pos, twf)
        plt.xlabel("Sample")
        plt.ylabel("Total Weight Fraction")
        plt.title("Total Weight Fractions")
        plt.xticks(x_pos, plants, rotation = 90)
        plt.tick_params(axis='x', which='major', labelsize=8)
        plt.tight_layout()

        # plt.show()
        plt.savefig(self.m_directory+'/total-weight-fractions.png')

    def GenerateGraphForProfiles(self):

        plants = self.GetPlants()
        columns = 5
        fig, axs = plt.subplots(math.ceil(len(plants)/columns), columns)
        fig.set_size_inches(100,100)
        for index, p in enumerate(plants):
            i = index-1
            print(f"Generating pie chart {i} of {len(plants)-1}; r={math.ceil(i/columns)} c={i % columns}")
            axs[math.floor(i/columns), i % columns].pie(
                [f.m_percentOfTotalFA for f in p.m_fames], 
                labels=[f.m_name for f in p.m_fames],
                autopct='%1.1f%%',
                shadow=False, 
                startangle=90,
                radius=2.5,
                frame=False,
                textprops={'fontsize': 8})
            # ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        plt.tight_layout(pad=4.0)
        plt.show()
