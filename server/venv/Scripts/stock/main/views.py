from django.shortcuts import render
from .models import Post
# Create your views here.
def index(request):
    return render(request,'main/index.html')

def blog(request):
    postlist = Post.objects.all()
    return render(request,'main/blog.html',{'postlist':postlist})