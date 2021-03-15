from sampleManagement import *
import matplotlib.pyplot as plt
import os
import math

class GraphGenerator:
    m_lineManager = None
    m_directory = "graphs"
    m_cachedSamples = []

    def __init__(self, lineManager, directory="graphs"):
        self.m_lineManager = lineManager
        self.m_directory = directory
        
        #Assume ctor is called from working directory or absolute path.
        if not os.path.exists(directory):
            os.makedirs(directory)

    def GetSamples(self):
        if (not len(self.m_cachedSamples)):
            self.m_cachedSamples = [s for l in self.m_lineManager.m_lines for s in l.m_samples]
        return self.m_cachedSamples

    def GenerateGraphForTotalWeightFraction(self):
        # plt.style.use('ggplot')
        samples = [s.m_name for s in self.GetSamples()]
        twf = [s.m_totalWeightFraction for s in self.GetSamples()]

        x_pos = [i for i, _ in enumerate(x)]

        plt.bar(x_pos, twf)
        plt.xlabel("Sample")
        plt.ylabel("Total Weight Fraction")
        plt.title("Total Weight Fractions")
        plt.xticks(x_pos, samples, rotation = 90)
        plt.tick_params(axis='x', which='major', labelsize=8)
        plt.tight_layout()

        # plt.show()
        plt.savefig(self.m_directory+'/total-weight-fractions.png')

    def GenerateGraphForProfiles(self):

        samples = self.GetSamples()
        columns = 5
        fig, axs = plt.subplots(math.ceil(len(samples)/columns), columns)
        fig.set_size_inches(100,100)
        for index, s in enumerate(samples):
            i = index-1
            print(f"Generating pie chart {i} of {len(samples)-1}; r={math.ceil(i/columns)} c={i % columns}")
            axs[math.floor(i/columns), i % columns].pie(
                [f.m_percentOfTotalFA for f in s.m_fames], 
                labels=[f.m_name for f in s.m_fames],
                autopct='%1.1f%%',
                shadow=False, 
                startangle=90,
                radius=2.5,
                frame=False,
                textprops={'fontsize': 8})
            # ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        plt.tight_layout(pad=4.0)
        plt.show()
